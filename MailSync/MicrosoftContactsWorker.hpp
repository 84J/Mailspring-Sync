//
//  MicrosoftContactsWorker.hpp
//  MailSync
//
//  Syncs the contacts of an Outlook.com / Microsoft 365 account through Microsoft Graph.
//
//  Microsoft serves no CardDAV, so the client otherwise only has found-in-mail contacts for
//  these accounts. Contacts are stored the way CardDAV contacts are - source "carddav" with
//  the card in info.vcf - because that is the format the client reads and edits. The Graph id
//  is kept in info.graphId and the Graph changeKey in the contact's etag.
//
//  Use of this file is subject to the terms and conditions defined
//  in 'LICENSE.md', which is part of the Mailspring-Sync package.
//

#ifndef MicrosoftContactsWorker_hpp
#define MicrosoftContactsWorker_hpp

#include "Account.hpp"
#include "Contact.hpp"
#include "ContactBook.hpp"
#include "MailStore.hpp"

#include <functional>
#include <memory>
#include <string>
#include <vector>

#include "json.hpp"
#include "spdlog/spdlog.h"

using namespace nlohmann;
using namespace std;

static string MSGRAPH_CONTACTS_BOOK_SOURCE = "msgraph";

class MicrosoftContactsWorker {
    unique_ptr<MailStore> store;
    shared_ptr<spdlog::logger> logger;
    shared_ptr<Account> account;

    std::function<string()> tokenProvider();
    shared_ptr<ContactBook> ensureBook();
    vector<string> listFolderURLs();
    void applyRemote(shared_ptr<Contact> local, const json & remote, const string & bookId);
    void runPass();

public:
    MicrosoftContactsWorker(shared_ptr<Account> account);

    // Accounts whose contacts this worker syncs: the same Microsoft accounts whose calendars
    // sync through Graph (a shared mailbox signs in as its owner, so /me would be the owner's).
    static bool appliesTo(shared_ptr<Account> account);

    // One full pass. Never throws: a failure is logged and the next pass tries again, so
    // contact sync can't stop calendar sync, which runs in the same loop.
    void run();

    // Writes the client's edit of `contact` (its info.vcf) to Graph, creating the contact when
    // it has no Graph id yet, then stores Graph's copy. Throws SyncException on failure.
    void upsertContact(shared_ptr<Contact> contact);

    // Deletes the contact from Graph and locally. Throws SyncException on failure.
    void deleteContact(shared_ptr<Contact> contact);
};

#endif /* MicrosoftContactsWorker_hpp */
