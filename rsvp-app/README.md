# RSVP setup and deployment

The RSVP form is a private Google Apps Script web app embedded in the public
wedding website. Guest names and responses stay in a private Google Sheet; do
not add the guest list, invitation codes, or the Sheet to this repository.

## 1. Create the private Sheet and script

1. Create a new Google Sheet in the Google account that will own the RSVP data.
2. In the Sheet, choose **Extensions → Apps Script**.
3. Replace the default script with the contents of `Code.gs`.
4. Add an HTML file named `Index` and paste in `Index.html`.
5. In **Project Settings**, enable **Show "appsscript.json" manifest file in
   editor**, then replace the manifest with `appsscript.json`.
6. Save the Apps Script project.

## 2. Create the RSVP worksheets

1. Reload the Google Sheet so the **RSVP Admin** menu appears.
2. Choose **RSVP Admin → 1. Set up RSVP sheets**.
3. Approve the Google authorization prompt.
4. Delete the two orange example rows in **Guest Setup**.

The setup sheet uses one row per guest. Every event column accepts `TRUE` or
`FALSE`. Events that collect an RSVP appear when marked `TRUE`; schedule-only
flags remain private for building each guest's personalized itinerary:

| Column | Meaning |
| --- | --- |
| `household_key` | An admin label shared by one household, such as `SMITH` |
| `household_name` | Greeting shown on the RSVP, such as `The Smith Family` |
| `invite_code` | A unique 2–40 character, space-free code shared by the household, such as `SmithSunset` |
| `contact_email` | Optional household email address |
| `guest_name` | Guest name shown on the RSVP |
| `rehearsal_invited` | Schedule-only flag for Rehearsal; not shown as an RSVP question |
| `ranta_reception_invited` | Show Ranta Reception |
| `welcome_celebration_invited` | Show Welcome Celebration |
| `bridal_party_prep_invited` | Schedule-only flag for Bridal Party Prep; not shown as an RSVP question |
| `pre_wedding_photos_invited` | Schedule-only flag for Pre-wedding Photos; not shown as an RSVP question |
| `wedding_invited` | Show Ceremony and Reception |
| `afterparty_invited` | Show Afterparty |
| `cabana_bay_sendoff_invited` | Show Cabana Bay Floatoff |

Rows with the same `household_key` must use the same household name, invite
code, and email. Codes may contain letters, numbers, hyphens, and underscores;
they are case-insensitive and must be unique across households. A last name is
easy to guess, so prefer an inside-joke phrase or a last name plus another word.
Each attendee's event answers and schedule-only invitation flags are stored on
that attendee's row in **Guests**.
The shared weekend details are asked once per family or party and stored on its
row in **Households**.

## 3. Generate invitations

1. Add a fake test household and its custom `invite_code` before loading the real guest list.
2. Choose **RSVP Admin → 2. Build invitations from Guest Setup**.
3. Keep **Households**, **Guests**, and especially **Invite Links** private.

Building invitations is intentionally allowed only once on a fresh Sheet so an
accidental rerun cannot erase responses. Use **RSVP Admin → Change one household
code** to choose a replacement if a code is exposed or sent to the wrong person.

## 4. Deploy the web app

1. In Apps Script, choose **Deploy → New deployment**.
2. Select **Web app**.
3. Set **Execute as** to **Me**.
4. Set **Who has access** to **Anyone**. The custom invitation code controls
   access to each household; guests should not need a Google account.
5. Deploy and copy the URL ending in `/exec`.

When the script changes later, create a new deployment version from **Deploy →
Manage deployments**. Keep using the same `/exec` URL on the wedding site.

## 5. Connect the wedding website

In the public `index.html`, replace the value of `rsvpAppEndpoint` with the
deployed `/exec` URL. Until then, the public page safely displays “Online RSVP
is being prepared.”

## 6. Test before using real invitations

Using the fake household's personalized link from **Invite Links**:

1. Give two fake attendees different combinations of RSVP and schedule-only event invitation flags.
2. Confirm the RSVP displays guests down the left and answerable events across the top.
3. Confirm each invited guest/event intersection has Yes/No controls and uninvited intersections show a dash.
4. Select Yes/No answers for every invited guest and answerable event.
5. Choose All of us, Some of Us, or None of Us for the Orlando question, and complete the Saturday ceremony arrival question once for the party.
6. Enter one shared set of song, karaoke, hotel, accessibility, and dietary information.
7. Save the RSVP.
8. Confirm event answers appear on the correct attendee rows in **Guests**.
9. Confirm the shared weekend answers appear on the correct row in **Households**.
10. Confirm **Your Weekend Schedule** appears above the RSVP matrix after saving.
11. Confirm it includes every event that at least one party member is invited to, including schedule-only events, with each date shown once above that day's events.
12. Confirm events limited to part of the party show the applicable guest names followed by “only.”
13. Reload the same link and confirm all saved answers and the personalized schedule return.
14. Change answers, save again, and confirm rows are updated, not duplicated.
15. Test the schedule and horizontally scrollable matrix on both a phone and a computer.

If a Sheet was created with an earlier RSVP schema—including one that stored
weekend questions per attendee—create a fresh Sheet before testing this version.

After the test passes, create a fresh production Sheet/script deployment for the
real guest list, or clear all test-generated data before building invitations.
