import requests, json, sqlite3, smtplib, smtplib, email, re, subprocess, os
from bs4 import BeautifulSoup


class ScrapeBuilds:
    def __init__(self):
        pass

    @staticmethod
    def fetch_office_insider_builds() -> dict:
        url = 'https://learn.microsoft.com/en-us/officeupdates/beta-channel'
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        title_heading = soup.find_all('h2')
        version_pattern = re.compile(r'Version .*:')

        office_insider = {}

        for title in title_heading:
            if version_pattern.search(title.text):
                key = title.text
                cleaning_url = key.lower().replace(' ', '-').split(':')
                clean_url = ''.join(cleaning_url)
                cleaning = str(title.find_next('em')).removeprefix('<em>').removesuffix('</em>').split('n ')

                data = {
                    'build': cleaning[1],
                    'url': f'{url}#{clean_url}'
                }

                office_insider[key] = data

        first_3_results = {}

        for key, value in list(office_insider.items())[:3]:
            first_3_results[key] = value

        return first_3_results

    @staticmethod
    def fetch_office_builds() -> dict:
        url = 'https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date'
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        table = soup.find('table')
        rows = table.find_all('tr')[1:]

        office = {}
        for row in rows:
            cells = row.find_all('td')
            data = {
                'build': cells[2].get_text(),
                'released': cells[3].get_text(),
            }
            key = f'{cells[0].get_text()} {cells[1].get_text()}'
            clean_date = cells[3].get_text().replace(" ", "-").split(",")
            amend_for_url = f'version-{cells[1].get_text()}-{clean_date[0]}'.lower()

            if "Current Channel" in key:
                data['url'] = f'https://learn.microsoft.com/en-us/officeupdates/' \
                              f'current-channel#{amend_for_url}'

            if "Monthly Enterprise Channel" in key:
                data['url'] = f'https://learn.microsoft.com/en-us/officeupdates/' \
                              f'monthly-enterprise-channel#{amend_for_url}'

            if "Semi-Annual Enterprise Channel" in key:
                data['url'] = f'https://learn.microsoft.com/en-us/officeupdates/' \
                              f'semi-annual-enterprise-channel#{amend_for_url}'

            if "Semi-Annual Enterprise Channel (Preview)" in key:
                data['url'] = f'https://learn.microsoft.com/en-us/officeupdates/' \
                              f'semi-annual-enterprise-channel-preview#{amend_for_url}'

            office[key] = data

        return office

    @staticmethod
    def fetch_teams_builds() -> dict:
        url = 'https://learn.microsoft.com/en-us/officeupdates/teams-app-versioning'
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        title_heading = soup.find('h3', string='Windows (GCC) version history')

        teams = {}
        if title_heading:
            table = title_heading.find_next_sibling('table')

            if table:
                rows = table.find_all('tr')[1:]

                for row in rows:
                    cells = row.find_all('td')
                    key = f'Version {cells[1].get_text()} {cells[0].get_text()}'
                    data = {
                        'build': cells[2].get_text(),
                    }

                    teams[key] = data

                first_3_results = {}

                for key, value in list(teams.items())[:3]:
                    first_3_results[key] = value

                return first_3_results

    @staticmethod
    def fetch_firefox_builds() -> dict:
        url = 'https://www.mozilla.org/en-US/firefox/releases/'
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        releases = soup.find_all('ol')
        dirty = str(releases).split('a href')

        firefox = {}

        for clean in dirty:
            if 'releasenotes' in clean:
                cleaning = clean.removeprefix('="../').split('/')
                cleaned = cleaning[0]
                major_version = cleaned.split('.')[0]
                key = f"Version {major_version}"

                if key not in firefox:
                    firefox[key] = []

                build_url = f'https://www.mozilla.org/en-US/firefox/{cleaned}/releasenotes/'
                firefox[key].append({'build': cleaned, 'url': build_url})

        for build_key in firefox:
            firefox[build_key] = sorted(firefox[build_key], key=lambda x: x['build'], reverse=True)[:5]

        show_results = {}
        for key, value in list(firefox.items())[:1]:
            show_results[key] = value

        return show_results

    @staticmethod
    def fetch_edge_builds() -> dict:
        url = 'https://learn.microsoft.com/en-us/deployedge/microsoft-edge-relnote-beta-channel'
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        title_heading = soup.find_all('h2')

        edge = {}

        for title in title_heading:
            dirty = str(title).split('id="')

            if 'Version' in dirty[1]:
                cleaning = dirty[1].removesuffix('</h2>').split('>')
                amend_for_url = cleaning[0].removesuffix('"')
                cleaned = cleaning[1].split(':')
                build = cleaned[0]
                date = cleaned[1]
                key = f'Version {date}'

                data = {
                    'build': build,
                    'url': f'{url}#{amend_for_url}'
                }

                edge[key] = data

        first_3_results = {}

        for key, value in list(edge.items())[:3]:
            first_3_results[key] = value

        return first_3_results

    @staticmethod
    def fetch_chrome_builds():
        url = 'https://chromereleases.googleblog.com/search/label/Stable%20updates'
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        title_heading = soup.find_all('h2')

        chrome = {}

        for url in title_heading:
            dirty = str(url).split('href="')
            try:
                cleaning = dirty[1].split('"')
                cleaning_date = str(url.find_next(class_='publishdate')).split('>')

                clean_date = cleaning_date[1].removesuffix('</span').strip()
                clean_url = cleaning[0]
                clean_title = cleaning[4]

                if 'Desktop' in clean_title:
                    key = clean_date
                    data = {
                        'channel': clean_title,
                        'url': clean_url,
                    }

                    chrome[key] = data

            except IndexError:
                # Handling any H2 headers that do not contain title tags
                continue

        first_3_results = {}

        for key, value in list(chrome.items())[:3]:
            first_3_results[key] = value

        return first_3_results

    @staticmethod
    def fetch_firefox_beta_builds() -> dict:
        url = 'https://www.mozilla.org/en-US/firefox/beta/notes'
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        find_class = soup.find_all(class_='c-release-version')
        find_date = soup.find_all(class_='c-release-date')

        version = str(find_class).split('>')
        date = str(find_date).split('>')

        key = version[1].removesuffix('beta</div')

        firefox_beta = {}
        data = {
            'date': date[1].removesuffix('</p'),
            'url': url
        }

        firefox_beta[key] = data
        return firefox_beta

    @staticmethod
    def fetch_thunderbird_beta_builds() -> dict:
        url = 'https://www.thunderbird.net/en-US/thunderbird/releases/'
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        releases = soup.find_all(
            'a', class_='inline-link text-blue-dark font-bold no-underline btn-body btn-secondary btn-release'
        )

        dirty = str(releases).split('href')
        thunderbird = {}

        for clean in dirty:
            if 'releasenotes/">' in clean:
                try:
                    cleaning = clean.split('="../../../en-US/thunderbird/')
                    cleaning2 = cleaning[1].split('/')
                    cleaned = cleaning2[0]

                except IndexError:
                    continue

                major_version = cleaned.split('.')
                key = f"Version {major_version[0]}"

                if key not in thunderbird:
                    thunderbird[key] = []

                build_url = f'https://www.thunderbird.net/en-US/thunderbird/{cleaned}/releasenotes/'
                thunderbird[key].append({'build': cleaned, 'url': build_url})

        for build_key in thunderbird:
            thunderbird[build_key] = sorted(thunderbird[build_key], key=lambda x: x['build'], reverse=True)[:5]

        show_results = {}
        for key, value in list(thunderbird.items())[:1]:
            show_results[key] = value

        return show_results


def content_format(key, value):
    content = f'A new update was located for {key}, ' \
              f'Please schedule testing on this build for the relevant products. \n\n' \
              f'See more information about the identified updates below:\n\n'

    update_keys = ["channel", "version", "build", "url", "released", "date"]

    for update_key, update_data in value.items():
        content += f'{update_key}\n'
        for key in update_keys:
            try:
                if key in update_data[0]:
                    content += f'{key.title()}: {update_data[0][key]}\n'
            except KeyError:
                if key in update_data:
                    content += f'{key.title()}: {update_data[key]}\n'
        content += '\n'
    return content


def send_email(sender_email: str, sender_pw: str, subject: str, send_to: list, content: str) -> bool:
    msg = email.message_from_string(content)
    msg['From'] = sender_email
    msg['To'] = ", ".join(send_to)
    msg['Subject'] = subject

    s = smtplib.SMTP("smtp-mail.outlook.com", 587)
    s.ehlo()
    s.starttls()
    s.ehlo()
    s.login(sender_email, sender_pw)

    try:
        s.sendmail(msg['From'], recipients, msg.as_string())
        return True

    except smtplib.SMTPException:
        s.quit()
        return False


def update_db() -> None:
    cur.execute("CREATE TABLE IF NOT EXISTS checker (key TEXT PRIMARY KEY, value TEXT)")

    for key, value in version_data.items():
        json_value = json.dumps(value, indent=3)
        cur.execute("SELECT * FROM checker WHERE key=?", (key,))
        existing_data = cur.fetchone()

        if existing_data:
            existing_json_value = existing_data[1]

            if json_value == existing_json_value:
                continue

            else:
                cur.execute("UPDATE checker SET value=? WHERE key=?", (json_value, key))
                send_email(
                    sender_email=sender_username,
                    sender_pw=sender_password,
                    subject=f'Test Coverage Checker - New {key} Update Available.',
                    send_to=recipients,
                    content=content_format(key, value)
                )
        else:
            cur.execute("INSERT INTO checker VALUES (?, ?)", (key, json_value))

    con.commit()
    con.close()


if __name__ == '__main__':
    user = os.environ['USERPROFILE']
    sender_username = 'web@yourdolphin.com'
    sender_password = 'Dolphin!986'

    recipients = [
        'betacontrol@yourdolphin.com',
        'joshua.garland@yourdolphin.com',
        'joshua.murphy@yourdolphin.com',
    ]

    version_data = {
        'Office 365 Stable': ScrapeBuilds.fetch_office_builds(),
        'Office 365 Insider': ScrapeBuilds.fetch_office_insider_builds(),
        'Teams': ScrapeBuilds.fetch_teams_builds(),
        'Edge': ScrapeBuilds.fetch_edge_builds(),
        'Firefox Release': ScrapeBuilds.fetch_firefox_builds(),
        'Firefox Beta': ScrapeBuilds.fetch_firefox_beta_builds(),
        'Chrome': ScrapeBuilds.fetch_chrome_builds(),
        'Thunderbird': ScrapeBuilds.fetch_thunderbird_beta_builds(),
    }

    con = sqlite3.connect(f"{user}/AppData/Roaming/checker.db")
    cur = con.cursor()
    update_db()
