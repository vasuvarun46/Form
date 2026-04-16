EformDetails (SPFx) — Full Page Setup
Prerequisites

Node.js: v18.20.8

npm: comes bundled with Node.js

⚠️ This project is sensitive to Node.js versions due to SPFx and Gulp tooling.
Use the exact Node version specified to avoid build/runtime issues.

Recommended Node Version Management

A .nvmrc file is provided at the project root to ensure consistent Node.js usage across environments.

18.20.8


If using nvm:

nvm install
nvm use

1) Global installs

Install required global tools:

npm install -g gulp-cli@2 yo@4


Verify installation:

gulp -v
yo --version


Install SPFx Yeoman generator (aligned with this project):

npm install -g @microsoft/generator-sharepoint@1.20.1

2) Project install

From the project root:

npm install


If you encounter dependency or peer warnings, ensure:

Node version is 18.20.8

npm cache is clean if required

3) Additional libraries used in this project
TailwindCSS (Current Setup – v3)

⚠️ Important Technical Note
This project uses TailwindCSS v3 integrated into the SPFx + Gulp pipeline.
Due to compatibility constraints between:

SPFx v1.20.x

Gulp build system

PostCSS pipeline

gulp serve does not function correctly with Tailwind enabled.
Therefore, gulp serve-deprecated is intentionally used for local development.

Installed versions:

npm install -D tailwindcss@3.4.17 \
postcss@8.4.38 \
autoprefixer@10.4.19 \
gulp-postcss \
postcss-import


This configuration is stable and verified for this codebase.

PDF Export (HTML → Image → PDF)
npm install html-to-image@1.11.13 jspdf@2.5.1

Excel Export
npm install xlsx@0.18.5

4) Local Development (Workbench)
Trust Developer Certificate (Required once per machine)
gulp trust-dev-cert

Run local workbench
gulp serve-deprecated


ℹ️ Why serve-deprecated is used

TailwindCSS relies on PostCSS processing

SPFx modern gulp serve has known limitations with custom PostCSS pipelines

gulp serve-deprecated provides a stable webpack/PostCSS flow for this setup

This is a deliberate technical choice, not a legacy oversight.

Notes / Troubleshooting
Clean install (recommended if issues occur)
rm -rf node_modules package-lock.json
npm install

Common fixes

Restart VS Code after installing Node

Ensure Node is available in PATH

Run commands from project root

Do not use Node 20+

Current Key Versions Used

SPFx tooling: @microsoft/sp-build-web@1.20.1

TypeScript: 4.7.4

React: 17.0.1

TailwindCSS: 3.4.17

Production Build & Packaging

Use the following steps for UAT / Production deployment:

gulp clean
gulp bundle --ship
gulp package-solution --ship


Generated package:

sharepoint/solution/*.sppkg


Upload the .sppkg file to:

Site Collection App Catalog

Deploy the app

Add it to the target SharePoint site

Required SharePoint Configuration

The solution assumes the following SharePoint lists exist:

Inventory Requests

Inventory Request Items

Approval History

Approval Matrix

Purpose Master

Cost Center Master

Requester Directory

Workflow Configuration

Project Overview – Inventory Request & Approval System (SPFx)

Role-Based Access & User Experience
Dynamically identifies users as Owner (Admin), Member (Approver), or Visitor (Requester) and adapts UI and actions accordingly.

Unified Inventory Request Creation
Supports manual entry and Excel upload in a single request form.

Draft Management & Auto-Recovery
Drafts can be saved, reopened, and edited before submission.

Multi-Level Approval Workflow
Approval routing is based on configurable business rules such as total amount.

Clear Role-Based Visibility

Visitors: own requests

Members: approval tasks

Owners: full visibility

Seamless Navigation
Smooth transitions across Home, Create Request, My Requests, Team Requests, and All Requests.

SharePoint & Power Automate Integration
SharePoint stores data; Power Automate handles notifications only.

Validation & Error Handling
Strong validations prevent incorrect submissions.

Export & Audit Support
Requests can be exported to PDF for audit and compliance.

Scalable & Maintainable Architecture
Built with SPFx, React, and TypeScript following modular principles.

Enterprise Quality Checklist (Client Review)

Before submitting code for review, ensure:

gulp serve-deprecated runs successfully

gulp bundle --ship succeeds

gulp package-solution --ship succeeds

Node version matches .nvmrc

Testing (Unit & UI Tests – SPFx Safe)
Testing Stack (SPFx Compatible)

The project uses a Jest-based testing setup aligned with SPFx 1.20.x and React 17.

Jest: 26.6.3

ts-jest: 26.5.6

@testing-library/react: 12.1.5

@testing-library/jest-dom: 5.16.5

Newer Jest versions (27+) are not compatible with SPFx due to jsdom/runtime constraints.

Install Testing Dependencies

From the project root:

npm install -D jest@26.6.3 ts-jest@26.5.6 jest-environment-jsdom@26.6.2 @types/jest@26.0.24
npm install -D @testing-library/react@12.1.5 @testing-library/jest-dom@5.16.5
npm install -D identity-obj-proxy

Test Scope

UI rendering tests for React components

Pure business and routing logic tests

SharePoint runtime, SPHttpClient, Power Automate, and permissions are intentionally excluded from unit tests and validated through integration testing.

Test Folder Structure
__mocks__/
├── fileMock.js
├── @microsoft/
│   ├── sp-core-library.js
│   ├── sp-http.js
│   └── sp-webpart-base.js

src/
├── tests/
│   ├── routeAccess.test.ts
│   ├── Sidebar.test.tsx
│   └── AppHeader.test.tsx

Run Tests
npm test
