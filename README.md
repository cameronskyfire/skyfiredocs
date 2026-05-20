# Skyfire Document Generator

Internal tool for generating completed Word documents (CIIAA, etc.) from templates.
Runs entirely in the browser — no backend, no data leaves the page.

## Usage

Open `index.html` (locally or via the hosted URL), enter the access password,
choose a document type, fill in the form, and download the completed `.docx`.

## Templates

Templates use `[[Placeholder Name]]` markers in Word. Add markers to your `.docx`
file, upload it via the "Add another template" section, and it will be saved
to your browser for repeat use. The CIIAA (Employee) template is pre-bundled.

## Changing the password

The access password is a SHA-256 hash stored in `index.html` near the bottom
of the file. To change it:

1. Open your browser's console (Cmd+Option+J or F12) on any page.
2. Run:
   ```js
   crypto.subtle.digest('SHA-256', new TextEncoder().encode('YourNewPassword'))
     .then(b => console.log([...new Uint8Array(b)].map(x => x.toString(16).padStart(2,'0')).join('')))
   ```
3. Copy the printed hash.
4. In `index.html`, find `const PASSWORD_HASH = '...'` and replace the value.
5. Commit and push — GitHub Pages will redeploy in a minute or two.

## Notes on security

Client-side password protection is a deterrent, not real security. The
template and password check live in the HTML served to the browser. For
proper access control, this would need a real backend.

