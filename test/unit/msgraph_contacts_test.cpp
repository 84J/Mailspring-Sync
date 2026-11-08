//
//  msgraph_contacts_test.cpp
//  Unit tests for the Microsoft Graph contact <-> vCard conversions.
//
//  Build and run (Linux): cmake . && make msgraph_contacts_test && ./msgraph_contacts_test
//

#include "MicrosoftContactsVCard.hpp"

#include <iostream>

using namespace MicrosoftContactsVCard;

static int failures = 0;

static void expectEqual(const json & actualValue, const string & expected, const string & label) {
    string actual = actualValue.is_string() ? actualValue.get<string>() : actualValue.dump();
    if (actual != expected) {
        cerr << "FAIL " << label << ": expected \"" << expected << "\", got \"" << actual << "\"\n";
        failures++;
    }
}

static void expectJSON(const json & actual, const json & expected, const string & label) {
    if (actual != expected) {
        cerr << "FAIL " << label << ": expected " << expected.dump() << ", got " << actual.dump() << "\n";
        failures++;
    }
}

static void expectContains(const string & haystack, const string & needle, const string & label) {
    if (haystack.find(needle) == string::npos) {
        cerr << "FAIL " << label << ": \"" << needle << "\" not found in:\n" << haystack << "\n";
        failures++;
    }
}

static const json SAMPLE = json::parse(R"({
    "id": "AAMkAD=",
    "changeKey": "EQAAABYA",
    "parentFolderId": "AAMkFolder=",
    "displayName": "Dr. Ada Lovelace",
    "givenName": "Ada",
    "middleName": "King",
    "surname": "Lovelace",
    "title": "Dr.",
    "generation": "",
    "nickName": "Ada",
    "companyName": "Analytical Engines; Ltd",
    "department": "Research",
    "jobTitle": "Programmer",
    "emailAddresses": [{"name": "Ada", "address": "ada@example.com"}, {"name": "Ada", "address": "ada@home.example"}],
    "businessPhones": ["+44 20 1234"],
    "homePhones": ["+44 20 5678"],
    "mobilePhone": "+44 7700 900",
    "homeAddress": {"street": "12 St James's Square", "city": "London", "state": "", "postalCode": "SW1Y 4JH", "countryOrRegion": "UK"},
    "businessAddress": {},
    "otherAddress": {"street": "", "city": "", "state": "", "postalCode": "", "countryOrRegion": ""},
    "birthday": "1815-12-10T11:59:00Z",
    "personalNotes": "First line, with comma\nSecond line",
    "businessHomePage": "https://example.com"
})");

int main() {
    // Graph -> vCard
    string card = vcardForGraphContact(SAMPLE);
    expectContains(card, "BEGIN:VCARD\r\nVERSION:3.0\r\n", "card header");
    expectContains(card, "UID:AAMkAD=\r\n", "uid is the graph id");
    expectContains(card, "FN:Dr. Ada Lovelace\r\n", "formatted name");
    expectContains(card, "N:Lovelace;Ada;King;Dr.;\r\n", "structured name");
    expectContains(card, "ORG:Analytical Engines\\; Ltd;Research\r\n", "org escapes semicolons");
    expectContains(card, "EMAIL;TYPE=internet:ada@example.com\r\n", "first email");
    expectContains(card, "TEL;TYPE=cell:+44 7700 900\r\n", "mobile phone");
    expectContains(card, "TEL;TYPE=work:+44 20 1234\r\n", "business phone");
    expectContains(card, "TEL;TYPE=home:+44 20 5678\r\n", "home phone");
    expectContains(card, "ADR;TYPE=home:;12 St James's Square;;London;;SW1Y 4JH;UK\r\n", "home address in client layout");
    expectContains(card, "BDAY:1815-12-10\r\n", "birthday as date");
    expectContains(card, "NOTE:First line\\, with comma\\nSecond line\r\n", "note escaping");
    if (card.find("ADR;TYPE=work") != string::npos || card.find("ADR;TYPE=other") != string::npos) {
        cerr << "FAIL empty addresses must be skipped\n";
        failures++;
    }

    // vCard -> Graph, round trip of our own card
    json fields = graphFieldsForVCard(card);
    expectEqual(fields["displayName"], "Dr. Ada Lovelace", "round trip displayName");
    expectEqual(fields["givenName"], "Ada", "round trip givenName");
    expectEqual(fields["middleName"], "King", "round trip middleName");
    expectEqual(fields["surname"], "Lovelace", "round trip surname");
    expectEqual(fields["title"], "Dr.", "round trip title");
    expectEqual(fields["companyName"], "Analytical Engines; Ltd", "round trip company");
    expectEqual(fields["department"], "Research", "round trip department");
    expectEqual(fields["jobTitle"], "Programmer", "round trip jobTitle");
    expectEqual(fields["mobilePhone"], "+44 7700 900", "round trip mobile");
    expectJSON(fields["businessPhones"], json::array({"+44 20 1234"}), "round trip business phones");
    expectJSON(fields["homePhones"], json::array({"+44 20 5678"}), "round trip home phones");
    expectEqual(fields["emailAddresses"][0]["address"], "ada@example.com", "round trip email 1");
    expectEqual(fields["emailAddresses"][1]["address"], "ada@home.example", "round trip email 2");
    expectEqual(fields["homeAddress"]["street"], "12 St James's Square", "round trip street");
    expectEqual(fields["homeAddress"]["postalCode"], "SW1Y 4JH", "round trip postal code");
    expectEqual(fields["birthday"], "1815-12-10T11:59:00Z", "round trip birthday");
    expectEqual(fields["personalNotes"], "First line, with comma\nSecond line", "round trip notes");
    expectEqual(fields["businessHomePage"], "https://example.com", "round trip url");

    // A card as the client's editor writes it: lower-case parameters, grouped lines, folded
    // lines, a removed phone, a v4-style birthday, and an address without a type.
    string edited =
        "BEGIN:VCARD\r\n"
        "VERSION:3.0\r\n"
        "UID:AAMkAD=\r\n"
        "N:Hopper;Grace;;;\r\n"
        "FN:Grace Hopper\r\n"
        "item1.EMAIL;type=INTERNET;type=pref:grace@navy.example\r\n"
        "TEL;type=CELL;type=VOICE:+1 555 0100\r\n"
        "ADR:;1 Navy\r\n"
        " Yard;;Arlington;VA;22202;USA\r\n"
        "BDAY:19061209\r\n"
        "END:VCARD\r\n";
    json g = graphFieldsForVCard(edited);
    expectEqual(g["displayName"], "Grace Hopper", "edited displayName");
    expectEqual(g["emailAddresses"][0]["address"], "grace@navy.example", "grouped email");
    expectEqual(g["emailAddresses"][0]["name"], "Grace Hopper", "email name from display name");
    expectEqual(g["mobilePhone"], "+1 555 0100", "cell with extra type");
    expectJSON(g["businessPhones"], json::array(), "removed business phone is cleared");
    expectJSON(g["homePhones"], json::array(), "removed home phone is cleared");
    expectEqual(g["businessAddress"]["street"], "1 NavyYard", "folded untyped address -> business");
    expectEqual(g["businessAddress"]["city"], "Arlington", "address city");
    expectEqual(g["homeAddress"]["street"], "", "absent home address is cleared");
    expectEqual(g["birthday"], "1906-12-09T11:59:00Z", "compact birthday");
    expectEqual(g["companyName"], "", "absent company is cleared");

    // A card without a birthday clears it in Graph
    json noBirthday = graphFieldsForVCard("BEGIN:VCARD\r\nVERSION:3.0\r\nFN:X\r\nEND:VCARD\r\n");
    if (!noBirthday["birthday"].is_null()) {
        cerr << "FAIL missing birthday must be sent as null\n";
        failures++;
    }

    // At most three email addresses
    string many = "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Many\r\n"
                  "EMAIL:a@x.example\r\nEMAIL:b@x.example\r\nEMAIL:c@x.example\r\nEMAIL:d@x.example\r\n"
                  "END:VCARD\r\n";
    expectEqual(to_string(graphFieldsForVCard(many)["emailAddresses"].size()), "3", "graph keeps three emails");

    if (failures) {
        cerr << failures << " failure(s)\n";
        return 1;
    }
    cout << "msgraph_contacts_test: all passed\n";
    return 0;
}
