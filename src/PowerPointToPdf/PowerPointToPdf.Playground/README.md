# PowerPointToPdf.Playground

React 19 playground for `.pptx` / `.ppt` files up to 10 MB.
Upload or drag a file, convert through `PowerPointToPdf.Proxy`, then view the PDF in a
new tab or download it. LibreOffice is the initial conversion engine.

```powershell
npm ci
npm run dev
```

Open http://localhost:5177. Set `POWERPOINT_TO_PDF_PROXY_URL` in `.env` (see `.env.sample`)
to change the development proxy target. Default: http://localhost:62833.
`npm run build` builds the frontend; `npm run test:e2e` runs browser tests.
See [the solution README](../README.md) for backend and Docker instructions.
