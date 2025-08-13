# Office Add-in Manifest Validation

This project includes a script to validate the Office add-in manifest using the official validator.

## Validate the manifest

```bash
npm run validate:manifest
```

This runs:

```bash
npx office-addin-manifest validate manifest-patched.xml
```

Or via the npm script defined in `package.json`:

```json
{
  "scripts": {
    "validate:manifest": "npx office-addin-manifest validate manifest-patched.xml"
  }
}
```

## Expected successful output

Below is an excerpt of a passing validation run from this repo:

```text
Validation Information:
Package Type Identified: Package of your add-in was parsed successfully.
Correct Package: Your package matches the submission type.
Valid Manifest Schema: Your manifest does adhere to the current set of XML schema definitions for Add-in manifests.
Manifest Version Correct Structure: The manifest version number has the correct structure for the platform that it supports.
Manifest Version Correct Value: The manifest version number is greater or equal to 1.0.
Manifest ID Valid Prefix: The product ID in the manifest has a valid prefix
  - Details: b24892eb-f900-5bb7-9c67-8f4b23e28e9d
Manifest ID Correct Structure: The structure of the product ID is correct.
  - Details: b24892eb-f900-5bb7-9c67-8f4b23e28e9d
Desktop Source Location Present: A desktop or default source location URL is found.
Secure Desktop Source Location: The manifest desktop source location URLs use HTTPS.
The manifest source location URLs are valid.: The manifest source location URLs are valid.
Supported Office Identified: Supported Office products were successfully determined.
Support URL Present: The manifest support URL is present.
  - Details: https://guavaexcel.netlify.app
Valid Support URL structure: The manifest support URL has valid structure.
Valid OnlineMeetingCommandSurface ExtensionPoint.: OnlineMeetingCommandSurface ExtensionPoint extracted from manifest is found to be valid.
High Resolution Icon Present: A high resolution icon element was expected and is present.
  - Details: https://guavaexcel.netlify.app/assets/icon-64.png
Supported High Resolution Icon URL File Extension: The manifest high resolution icon URL has a valid image file extension.
  - Details: png
Secure High Resolution Icon URL: The manifest high resolution icon URL uses HTTPS.
  - Details: https://guavaexcel.netlify.app/assets/icon-64.png
Icon Present: A icon element was expected and is present.
  - Details: https://guavaexcel.netlify.app/assets/icon-32.png
Supported Icon URL File Extension: The manifest icon URL has a valid image file extension.
  - Details: png
The manifest icon URL uses HTTPS.: Secure Icon URL
  - Details: https://guavaexcel.netlify.app/assets/icon-32.png
All GetStarted strings are present in Resources: All GetStarted strings are present in Resources
Acceptance Test Completed: Acceptance test service has finished checking provided add-in.

The manifest is valid.
```

## Notes
- The manifest under validation is `manifest-patched.xml` at the project root.
- Ensure Node.js is installed. The command uses `npx` to fetch `office-addin-manifest` if it is not already installed locally.
