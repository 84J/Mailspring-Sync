//
//  MicrosoftContactsVCard.hpp
//  MailSync
//
//  Conversions between Microsoft Graph contacts and the vCards the client reads and edits.
//
//  Use of this file is subject to the terms and conditions defined
//  in 'LICENSE.md', which is part of the Mailspring-Sync package.
//

#ifndef MicrosoftContactsVCard_hpp
#define MicrosoftContactsVCard_hpp

#include <string>

#include "json.hpp"

using namespace nlohmann;
using namespace std;

namespace MicrosoftContactsVCard {

// A vCard 3.0 for a Graph contact resource, in the shape the client's contact editor reads
// (ADR as ;street;extended;city;region;postalCode;country, BDAY as YYYY-MM-DD).
string vcardForGraphContact(const json & remote);

// The Graph contact fields for a card the client edited. Every mapped field is present (empty
// when the card lacks it) so a value removed in the client is removed in Outlook too.
json graphFieldsForVCard(const string & vcf);

// object[key] when it is a string, else "".
string jsonStr(const json & object, const string & key);

string trim(const string & s);

}

#endif /* MicrosoftContactsVCard_hpp */
