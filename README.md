# Coursing

A senior-phase coursing triage tool. Imports a SEEMiS allocation export
(.xlsx) and optional Insight `Total_Attainment` export, runs rule-based flags
(under-loaded, excess study, AH without Higher A/B, Higher with weak N5,
SEEMiS workflow markers), and produces a change list for the data manager.

All data lives client-side in IndexedDB.

## Run locally

```sh
npm install
npm run dev
```

Then open the printed local URL.

## Build

```sh
npm run build
npm run preview
```
