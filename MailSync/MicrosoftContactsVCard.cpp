//
//  MicrosoftContactsVCard.cpp
//  MailSync
//
//  Conversions between Microsoft Graph contacts and the vCards the client reads and edits.
//  No mailcore or database dependency, so it can be unit tested on its own.
//
//  Use of this file is subject to the terms and conditions defined
//  in 'LICENSE.md', which is part of the Mailspring-Sync package.
//

#include "MicrosoftContactsVCard.hpp"

#include <algorithm>
#include <cctype>
#include <functional>
#include <set>
#include <vector>

namespace MicrosoftContactsVCard {

// ---------------------------------------------------------------------------------------
// vCard helpers

static string vcardEscape(const string & value) {
    string out;
    for (char c : value) {
        switch (c) {
            case '\\': out += "\\\\"; break;
            case ',': out += "\\,"; break;
            case ';': out += "\\;"; break;
            case '\n': out += "\\n"; break;
            case '\r': break;
            default: out += c;
        }
    }
    return out;
}

static string vcardUnescape(const string & value) {
    string out;
    for (size_t i = 0; i < value.size(); i++) {
        if (value[i] == '\\' && i + 1 < value.size()) {
            char n = value[++i];
            out += (n == 'n' || n == 'N') ? '\n' : n;
        } else {
            out += value[i];
        }
    }
    return out;
}

// Splits a structured value (N, ADR, ORG) at unescaped semicolons, unescaping each part.
static vector<string> vcardComponents(const string & value) {
    vector<string> parts {""};
    for (size_t i = 0; i < value.size(); i++) {
        if (value[i] == '\\' && i + 1 < value.size()) {
            parts.back() += value[i];
            parts.back() += value[++i];
        } else if (value[i] == ';') {
            parts.push_back("");
        } else {
            parts.back() += value[i];
        }
    }
    for (auto & p : parts) {
        p = vcardUnescape(p);
    }
    return parts;
}

static string lower(string s) {
    transform(s.begin(), s.end(), s.begin(), [](unsigned char c) { return (char)tolower(c); });
    return s;
}

string trim(const string & s) {
    size_t b = s.find_first_not_of(" \t\r\n");
    if (b == string::npos) return "";
    size_t e = s.find_last_not_of(" \t\r\n");
    return s.substr(b, e - b + 1);
}

struct VCardLine {
    string name;          // upper case, without group prefix or parameters
    set<string> types;    // lower case TYPE values (TYPE=a,b and bare parameters)
    string value;         // raw (still escaped)
};

// Unfolds and splits a card into its content lines.
static vector<VCardLine> vcardLines(const string & vcf) {
    vector<string> raw;
    string current;
    size_t start = 0;
    while (start <= vcf.size()) {
        size_t end = vcf.find('\n', start);
        string line = vcf.substr(start, end == string::npos ? string::npos : end - start);
        if (!line.empty() && line.back() == '\r') line.pop_back();
        if (!line.empty() && (line[0] == ' ' || line[0] == '\t')) {
            current += line.substr(1);
        } else {
            if (!current.empty()) raw.push_back(current);
            current = line;
        }
        if (end == string::npos) break;
        start = end + 1;
    }
    if (!current.empty()) raw.push_back(current);

    vector<VCardLine> lines;
    for (auto & l : raw) {
        // the name/parameter part ends at the first colon outside a quoted parameter value
        bool quoted = false;
        size_t colon = string::npos;
        for (size_t i = 0; i < l.size(); i++) {
            if (l[i] == '"') quoted = !quoted;
            if (l[i] == ':' && !quoted) { colon = i; break; }
        }
        if (colon == string::npos) continue;

        VCardLine out;
        out.value = l.substr(colon + 1);
        string head = l.substr(0, colon);
        vector<string> params;
        size_t p = 0;
        while (true) {
            size_t semi = head.find(';', p);
            params.push_back(head.substr(p, semi == string::npos ? string::npos : semi - p));
            if (semi == string::npos) break;
            p = semi + 1;
        }
        string name = params[0];
        size_t dot = name.find('.');
        if (dot != string::npos) name = name.substr(dot + 1); // "item1.EMAIL"
        transform(name.begin(), name.end(), name.begin(), ::toupper);
        out.name = name;

        for (size_t i = 1; i < params.size(); i++) {
            string param = lower(params[i]);
            string values = param;
            size_t eq = param.find('=');
            if (eq != string::npos) {
                if (param.substr(0, eq) != "type") continue;
                values = param.substr(eq + 1);
            }
            values.erase(remove(values.begin(), values.end(), '"'), values.end());
            size_t q = 0;
            while (true) {
                size_t comma = values.find(',', q);
                string t = values.substr(q, comma == string::npos ? string::npos : comma - q);
                if (!t.empty()) out.types.insert(t);
                if (comma == string::npos) break;
                q = comma + 1;
            }
        }
        lines.push_back(out);
    }
    return lines;
}

string jsonStr(const json & object, const string & key) {
    if (object.is_object() && object.count(key) && object[key].is_string()) {
        return object[key].get<string>();
    }
    return "";
}

static bool addressIsEmpty(const json & a) {
    for (auto key : {"street", "city", "state", "postalCode", "countryOrRegion"}) {
        if (trim(jsonStr(a, key)) != "") return false;
    }
    return true;
}

// ---------------------------------------------------------------------------------------
// Graph contact <-> vCard

string vcardForGraphContact(const json & remote) {
    string out = "BEGIN:VCARD\r\nVERSION:3.0\r\n";
    auto line = [&](const string & name, const string & value) {
        out += name + ":" + value + "\r\n";
    };

    line("UID", vcardEscape(jsonStr(remote, "id")));

    string given = jsonStr(remote, "givenName");
    string middle = jsonStr(remote, "middleName");
    string surname = jsonStr(remote, "surname");
    string display = jsonStr(remote, "displayName");
    if (display == "") {
        display = trim(given + " " + surname);
    }
    line("FN", vcardEscape(display));
    line("N", vcardEscape(surname) + ";" + vcardEscape(given) + ";" + vcardEscape(middle) + ";" +
              vcardEscape(jsonStr(remote, "title")) + ";" + vcardEscape(jsonStr(remote, "generation")));

    if (jsonStr(remote, "nickName") != "") {
        line("NICKNAME", vcardEscape(jsonStr(remote, "nickName")));
    }
    if (jsonStr(remote, "companyName") != "" || jsonStr(remote, "department") != "") {
        string org = vcardEscape(jsonStr(remote, "companyName"));
        if (jsonStr(remote, "department") != "") {
            org += ";" + vcardEscape(jsonStr(remote, "department"));
        }
        line("ORG", org);
    }
    if (jsonStr(remote, "jobTitle") != "") {
        line("TITLE", vcardEscape(jsonStr(remote, "jobTitle")));
    }

    if (remote.count("emailAddresses") && remote["emailAddresses"].is_array()) {
        for (const auto & e : remote["emailAddresses"]) {
            string address = trim(jsonStr(e, "address"));
            if (address != "") {
                line("EMAIL;TYPE=internet", vcardEscape(address));
            }
        }
    }

    auto phones = [&](const string & key, const string & type) {
        if (remote.count(key) && remote[key].is_array()) {
            for (const auto & p : remote[key]) {
                if (p.is_string() && trim(p.get<string>()) != "") {
                    line("TEL;TYPE=" + type, vcardEscape(trim(p.get<string>())));
                }
            }
        }
    };
    if (trim(jsonStr(remote, "mobilePhone")) != "") {
        line("TEL;TYPE=cell", vcardEscape(trim(jsonStr(remote, "mobilePhone"))));
    }
    phones("businessPhones", "work");
    phones("homePhones", "home");

    // The client reads ADR as ;street;extended;city;region;postalCode;country.
    auto address = [&](const string & key, const string & type) {
        if (!remote.count(key) || !remote[key].is_object() || addressIsEmpty(remote[key])) {
            return;
        }
        const json & a = remote[key];
        line("ADR;TYPE=" + type, ";" + vcardEscape(jsonStr(a, "street")) + ";;" +
             vcardEscape(jsonStr(a, "city")) + ";" + vcardEscape(jsonStr(a, "state")) + ";" +
             vcardEscape(jsonStr(a, "postalCode")) + ";" + vcardEscape(jsonStr(a, "countryOrRegion")));
    };
    address("businessAddress", "work");
    address("homeAddress", "home");
    address("otherAddress", "other");

    // Graph gives a timestamp (Outlook stores birthdays at 11:59 UTC); the client wants YYYY-MM-DD.
    string birthday = jsonStr(remote, "birthday");
    if (birthday.size() >= 10) {
        line("BDAY", birthday.substr(0, 10));
    }
    if (trim(jsonStr(remote, "businessHomePage")) != "") {
        line("URL;TYPE=work", vcardEscape(trim(jsonStr(remote, "businessHomePage"))));
    }
    if (jsonStr(remote, "personalNotes") != "") {
        line("NOTE", vcardEscape(jsonStr(remote, "personalNotes")));
    }

    out += "END:VCARD\r\n";
    return out;
}

json graphFieldsForVCard(const string & vcf) {
    json emails = json::array();
    json businessPhones = json::array();
    json homePhones = json::array();
    string mobilePhone = "";
    json emptyAddress = {{"street", ""}, {"city", ""}, {"state", ""}, {"postalCode", ""}, {"countryOrRegion", ""}};
    json homeAddress = emptyAddress, businessAddress = emptyAddress, otherAddress = emptyAddress;
    bool haveHome = false, haveBusiness = false, haveOther = false;
    string displayName, given, middle, surname, prefix, suffix;
    string nickName, company, department, jobTitle, birthday, homePage, notes;

    for (auto & l : vcardLines(vcf)) {
        if (l.name == "FN") {
            displayName = vcardUnescape(l.value);
        } else if (l.name == "N") {
            auto parts = vcardComponents(l.value);
            parts.resize(5);
            surname = parts[0]; given = parts[1]; middle = parts[2]; prefix = parts[3]; suffix = parts[4];
        } else if (l.name == "NICKNAME" && nickName == "") {
            nickName = vcardComponents(l.value)[0];
        } else if (l.name == "ORG" && company == "") {
            auto parts = vcardComponents(l.value);
            company = parts[0];
            department = parts.size() > 1 ? parts[1] : "";
        } else if (l.name == "TITLE" && jobTitle == "") {
            jobTitle = vcardUnescape(l.value);
        } else if (l.name == "EMAIL") {
            string address = trim(vcardUnescape(l.value));
            // Graph stores at most three addresses per contact.
            if (address != "" && emails.size() < 3) {
                emails.push_back({{"address", address}});
            }
        } else if (l.name == "TEL") {
            string number = trim(vcardUnescape(l.value));
            if (number == "") continue;
            if ((l.types.count("cell") || l.types.count("mobile")) && mobilePhone == "") {
                mobilePhone = number;
            } else if (l.types.count("home")) {
                homePhones.push_back(number);
            } else {
                businessPhones.push_back(number);
            }
        } else if (l.name == "ADR") {
            auto parts = vcardComponents(l.value);
            parts.resize(7);
            string street = trim(parts[1]);
            if (trim(parts[2]) != "") {
                street = trim(street + "\n" + trim(parts[2]));
            }
            json a = {{"street", street}, {"city", parts[3]}, {"state", parts[4]},
                      {"postalCode", parts[5]}, {"countryOrRegion", parts[6]}};
            if (l.types.count("home") && !haveHome) {
                homeAddress = a; haveHome = true;
            } else if (l.types.count("other") && !haveOther) {
                otherAddress = a; haveOther = true;
            } else if (!haveBusiness) {
                businessAddress = a; haveBusiness = true;
            } else if (!haveOther) {
                otherAddress = a; haveOther = true;
            }
        } else if (l.name == "BDAY") {
            string v = trim(l.value);
            string digits;
            for (char c : v) {
                if (isdigit((unsigned char)c)) digits += c;
                if (c == 'T') break;
            }
            if (digits.size() >= 8) {
                birthday = digits.substr(0, 4) + "-" + digits.substr(4, 2) + "-" + digits.substr(6, 2) + "T11:59:00Z";
            }
        } else if (l.name == "URL" && homePage == "") {
            homePage = trim(vcardUnescape(l.value));
        } else if (l.name == "NOTE") {
            notes = vcardUnescape(l.value);
        }
    }

    if (displayName == "") {
        displayName = trim(given + " " + surname);
    }
    for (auto & e : emails) {
        e["name"] = displayName;
    }

    // Every mapped field is sent, empty when the card lacks it, so removing a value in the
    // client removes it in Outlook too.
    json fields = {
        {"displayName", displayName},
        {"givenName", given},
        {"middleName", middle},
        {"surname", surname},
        {"title", prefix},
        {"generation", suffix},
        {"nickName", nickName},
        {"companyName", company},
        {"department", department},
        {"jobTitle", jobTitle},
        {"emailAddresses", emails},
        {"businessPhones", businessPhones},
        {"homePhones", homePhones},
        {"mobilePhone", mobilePhone},
        {"homeAddress", homeAddress},
        {"businessAddress", businessAddress},
        {"otherAddress", otherAddress},
        {"businessHomePage", homePage},
        {"personalNotes", notes},
    };
    fields["birthday"] = birthday == "" ? json(nullptr) : json(birthday);
    return fields;
}


} // namespace MicrosoftContactsVCard
