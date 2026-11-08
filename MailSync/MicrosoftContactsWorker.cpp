//
//  MicrosoftContactsWorker.cpp
//  MailSync
//
//  Use of this file is subject to the terms and conditions defined
//  in 'LICENSE.md', which is part of the Mailspring-Sync package.
//

#include "MicrosoftContactsWorker.hpp"
#include "MicrosoftContactsVCard.hpp"
#include "DAVUtils.hpp"
#include "JSONUtils.hpp"
#include "MailStoreTransaction.hpp"
#include "MailUtils.hpp"
#include "NetworkRequestUtils.hpp"
#include "SyncException.hpp"
#include "XOAuth2TokenManager.hpp"

#include <algorithm>
#include <cctype>
#include <mutex>
#include <set>
#include <unordered_map>

namespace MCV = MicrosoftContactsVCard;

// The fields read from Graph and mapped to the card. Anything else on the Graph contact is
// left alone: writes PATCH only the fields below.
static const string CONTACT_FIELDS =
    "id,changeKey,parentFolderId,displayName,givenName,middleName,surname,title,generation,"
    "nickName,companyName,department,jobTitle,emailAddresses,businessPhones,homePhones,"
    "mobilePhone,homeAddress,businessAddress,otherAddress,birthday,personalNotes,businessHomePage";

// Contact folders nest; deeper ones are rare and are skipped beyond this depth.
static const int MAX_FOLDER_DEPTH = 4;

// One pass or write at a time per process: the background pass and the task processor each
// hold their own worker, but both change the same rows.
static std::mutex contactsMtx;

// ---------------------------------------------------------------------------------------
// Worker

MicrosoftContactsWorker::MicrosoftContactsWorker(shared_ptr<Account> account) :
    store(make_unique<MailStore>()),
    logger(spdlog::get("logger")),
    account(account)
{
}

bool MicrosoftContactsWorker::appliesTo(shared_ptr<Account> account) {
    return account->syncsCalendarsViaGraph();
}

std::function<string()> MicrosoftContactsWorker::tokenProvider() {
    return [this]() {
        return SharedXOAuth2TokenManager()->partsForAccount(account, OAuthTokenAudience::MicrosoftGraphContacts).accessToken;
    };
}

shared_ptr<ContactBook> MicrosoftContactsWorker::ensureBook() {
    // The contact book is how the client knows contact sync is enabled for the account.
    auto book = store->find<ContactBook>(Query().equal("accountId", account->id()));
    if (!book) {
        book = make_shared<ContactBook>(account->id() + "-default", account->id());
    }
    if (book->source() != MSGRAPH_CONTACTS_BOOK_SOURCE) {
        book->setSource(MSGRAPH_CONTACTS_BOOK_SOURCE);
        store->save(book.get());
    }
    return book;
}

// The contacts collection of the default folder, then of every folder below it.
vector<string> MicrosoftContactsWorker::listFolderURLs() {
    vector<string> urls { MicrosoftGraphURL() + "/me/contacts" };

    std::function<void(const string &, int)> visit = [&](const string & listURL, int depth) {
        vector<string> ids;
        PaginateGraphCollection(listURL + "?$select=id&$top=100", tokenProvider(), [&](json & page) {
            if (!page.count("value")) return;
            for (const auto & folder : page["value"]) {
                string id = MCV::jsonStr(folder, "id");
                if (id != "") ids.push_back(id);
            }
        });
        for (const auto & id : ids) {
            string folderURL = MicrosoftGraphURL() + "/me/contactFolders/" + URLEscape(id);
            urls.push_back(folderURL + "/contacts");
            if (depth < MAX_FOLDER_DEPTH) {
                visit(folderURL + "/childFolders", depth + 1);
            }
        }
    };
    visit(MicrosoftGraphURL() + "/me/contactFolders", 1);
    return urls;
}

void MicrosoftContactsWorker::applyRemote(shared_ptr<Contact> local, const json & remote, const string & bookId) {
    string display = MCV::jsonStr(remote, "displayName");
    if (display == "") {
        display = MCV::trim(MCV::jsonStr(remote, "givenName") + " " + MCV::jsonStr(remote, "surname"));
    }
    string email = "";
    if (remote.count("emailAddresses") && remote["emailAddresses"].is_array()) {
        for (const auto & e : remote["emailAddresses"]) {
            email = MCV::trim(MCV::jsonStr(e, "address"));
            if (email != "") break;
        }
    }

    json info = local->info().is_object() ? local->info() : json::object();
    info["vcf"] = MCV::vcardForGraphContact(remote);
    info["graphId"] = MCV::jsonStr(remote, "id");
    info["graphFolderId"] = MCV::jsonStr(remote, "parentFolderId");

    local->setInfo(info);
    local->setName(display);
    local->setEmail(email);
    local->setEtag(MCV::jsonStr(remote, "changeKey"));
    local->setBookId(bookId);
}

void MicrosoftContactsWorker::run() {
    try {
        runPass();
    } catch (SyncException & ex) {
        logger->warn("-X Microsoft contacts could not be synced: {} {}", ex.key, ex.debuginfo);
    } catch (std::exception & ex) {
        logger->warn("-X Microsoft contacts could not be synced: {}", ex.what());
    }
}

void MicrosoftContactsWorker::runPass() {
    lock_guard<mutex> pass(contactsMtx);
    auto book = ensureBook();

    // Contacts already synced, by Graph id. Ones the client created that have no Graph id yet
    // are waiting for their syncback task and are left alone.
    unordered_map<string, shared_ptr<Contact>> byGraphId;
    for (auto & c : store->findAll<Contact>(Query().equal("accountId", account->id()).equal("source", CARDDAV_SYNC_SOURCE))) {
        json info = c->info();
        string graphId = MCV::jsonStr(info, "graphId");
        if (graphId != "") {
            byGraphId[graphId] = c;
        }
    }

    // Fetch everything first: a pass that fails halfway must not delete what it didn't reach.
    vector<json> remote;
    for (const auto & url : listFolderURLs()) {
        PaginateGraphCollection(url + "?$select=" + CONTACT_FIELDS + "&$top=500", tokenProvider(), [&](json & page) {
            if (!page.count("value")) return;
            for (auto & contact : page["value"]) {
                if (MCV::jsonStr(contact, "id") != "") remote.push_back(contact);
            }
        });
    }

    size_t changed = 0, removed = 0;
    set<string> seen;
    {
        MailStoreTransaction transaction{store.get(), "microsoftContacts"};
        for (const auto & r : remote) {
            string graphId = MCV::jsonStr(r, "id");
            seen.insert(graphId);
            auto it = byGraphId.find(graphId);
            shared_ptr<Contact> local = it != byGraphId.end() ? it->second : nullptr;
            if (local && local->etag() == MCV::jsonStr(r, "changeKey") && local->bookId() == book->id()) {
                continue;
            }
            if (!local) {
                local = make_shared<Contact>(MailUtils::idRandomlyGenerated(), account->id(), "", CONTACT_MAX_REFS, CARDDAV_SYNC_SOURCE);
            }
            applyRemote(local, r, book->id());
            store->save(local.get());
            changed++;
        }
        for (auto & pair : byGraphId) {
            if (!seen.count(pair.first)) {
                store->remove(pair.second.get());
                removed++;
            }
        }
        transaction.commit();
    }
    logger->info("Microsoft contacts: {} on server, {} updated, {} removed", remote.size(), changed, removed);
}

void MicrosoftContactsWorker::upsertContact(shared_ptr<Contact> contact) {
    lock_guard<mutex> pass(contactsMtx);
    auto book = ensureBook();
    json info = contact->info();
    string vcf = MCV::jsonStr(info, "vcf");
    if (vcf == "") {
        throw SyncException("invalid-contact", "Contact has no vCard to write to Microsoft Graph", false);
    }
    json fields = MCV::graphFieldsForVCard(vcf);
    string graphId = MCV::jsonStr(info, "graphId");

    GraphResponse resp;
    string url;
    if (graphId == "") {
        url = MicrosoftGraphURL() + "/me/contacts";
        logger->info("Creating Microsoft contact {}", contact->id());
        resp = PerformGraphRequest("POST", url, tokenProvider()(), fields.dump());
    } else {
        url = MicrosoftGraphURL() + "/me/contacts/" + URLEscape(graphId);
        logger->info("Updating Microsoft contact {}", contact->id());
        resp = PerformGraphRequest("PATCH", url, tokenProvider()(), fields.dump());
    }
    if (!resp.ok()) {
        ThrowForGraphResponse(url, resp);
    }

    // Read Graph's copy back: it fills in fields (displayName, changeKey) the write left out.
    string id = MCV::jsonStr(resp.body, "id");
    if (id == "") id = graphId;
    string readURL = MicrosoftGraphURL() + "/me/contacts/" + URLEscape(id) + "?$select=" + CONTACT_FIELDS;
    GraphResponse read = PerformGraphRequest("GET", readURL, tokenProvider()());
    if (!read.ok()) {
        ThrowForGraphResponse(readURL, read);
    }
    applyRemote(contact, read.body, book->id());
    store->save(contact.get());
}

void MicrosoftContactsWorker::deleteContact(shared_ptr<Contact> contact) {
    lock_guard<mutex> pass(contactsMtx);
    string graphId = MCV::jsonStr(contact->info(), "graphId");
    if (graphId != "") {
        string url = MicrosoftGraphURL() + "/me/contacts/" + URLEscape(graphId);
        logger->info("Deleting Microsoft contact {}", contact->id());
        GraphResponse resp = PerformGraphRequest("DELETE", url, tokenProvider()());
        // 404: already gone from Outlook, which is what we wanted.
        if (!resp.ok() && resp.status != 404) {
            ThrowForGraphResponse(url, resp);
        }
    }
    store->remove(contact.get());
}
