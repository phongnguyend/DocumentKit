# PowerPoint to PDF

The same CLI / API / proxy / React 19 playground structure as `HtmlToPdf`,
using LibreOffice Impress for `.pptx` and `.ppt` files, or the native
`PptxObjectKit.Extensions.Pdf` converter for `.pptx` files. Select either engine
or both in the playground to compare their PDFs.

| Project | Purpose |
| --- | --- |
| `PowerPointToPdf.LibreOffice.Cli` | Convert a local file from the command line |
| `PowerPointToPdf.LibreOffice.Api` | Multipart file conversion API |
| `PowerPointToPdf.PptxObjectKit.Api` | Native .NET PPTX-to-PDF engine (no LibreOffice required) |
| `PowerPointToPdf.Proxy` | YARP gateway at `/libreoffice/*` and `/pptxobjectkit/*` |
| `PowerPointToPdf.Playground` | Drag/drop or browse, then view or download the PDF |

LibreOffice execution lives in the API project and is linked into its CLI. Excel, Word and PowerPoint have independent converter implementations. Each conversion has an
isolated temporary directory and LibreOffice profile, macros disabled, a
90-second timeout, process-tree termination on cancellation, and cleanup.

## Docker Compose

To use released Docker Hub images, follow the [published-image setup](../../README.md#run-published-docker-images)
and run `docker compose -f compose.published.yaml up -d` after configuring
`DOCKERHUB_NAMESPACE` and `IMAGE_TAG` in this folder's `.env` file.
Copy `.env.sample` to `.env` if needed; `IMAGE_TAG` defaults to `latest`.
The commands below build images locally.

From this directory:

```powershell
docker compose up --build -d
```

| Service | URL |
| --- | --- |
| Playground | http://localhost:8114 |
| Proxy | http://localhost:8110 |
| LibreOffice API | http://localhost:8111 |
| Native API | http://localhost:8112 |

The native API Dockerfile uses `src/` as its build context to include PptxObjectKit. Other .NET Dockerfiles use this solution directory.
Their `Dockerfile.dockerignore` excludes local `bin/` and `obj/` output.
Containers run as a non-root user. The native engine includes fonts and Skia, without LibreOffice.

## Local development

Requires .NET 11 SDK, Node.js 22.12+, npm, and LibreOffice.
Set `LIBREOFFICE_BIN` if LibreOffice is not at the standard Windows location
or available as `soffice` on PATH on Linux. Start these in separate terminals:

```powershell
dotnet build ./PowerPointToPdf.slnx
dotnet run --project ./PowerPointToPdf.LibreOffice.Api
dotnet run --project ./PowerPointToPdf.PptxObjectKit.Api
dotnet run --project ./PowerPointToPdf.Proxy
```

In `PowerPointToPdf.Playground`, run `npm ci` and `npm run dev`, then open
http://localhost:5177. Vite forwards `/api/*` to the proxy at
http://localhost:62833, removing `/api`.
Copy `.env.sample` to `.env` to configure `POWERPOINT_TO_PDF_PROXY_URL`; use
`http://localhost:8110` for a Compose backend. Restart Vite after edits.
Docker uses Nginx forwarding to the proxy service, not the Vite environment variable.

## API

`POST /libreoffice/convert` on the proxy accepts `multipart/form-data` field
`file`, with a non-empty `.pptx` or `.ppt` upload up to 10 MB.
The API itself exposes `POST /convert` and `GET /health`; the proxy also has
`GET /health`. Success returns `application/pdf` with a download filename.
Invalid uploads return 400, unconvertible documents 422, missing LibreOffice
503, and conversion timeouts 504. Password-protected documents are unsupported.

```powershell
curl.exe -f -F "file=@sample.pptx" http://localhost:8110/libreoffice/convert -o sample.pdf
```

## Native engine

`POST /pptxobjectkit/convert` on the proxy accepts the same multipart `file`
field, limited to non-empty `.pptx` files up to 10 MB. The service runs locally
on port 62831. It returns PDF bytes and an `X-Conversion-Warning-Count` header;
the playground displays a notice when content was approximated or omitted.
Invalid uploads return 400 and unsupported or malformed presentations return 422.

This engine preserves supported text as selectable text and supports images and
vector graphics. It is experimental: complex Office graphics, effects and some
fonts/layouts may differ. Compare its output with LibreOffice before sharing.
The implementation is `PptxObjectKit.Extensions.Pdf`; `PdfObjectKit.Extensions.Pptx`
converts PDF to PowerPoint in the opposite direction.

```powershell
curl.exe -f -F "file=@sample.pptx" http://localhost:8110/pptxobjectkit/convert -o native.pdf
```

## CLI

```powershell
dotnet run --project ./PowerPointToPdf.LibreOffice.Cli -- input.pptx output.pdf
```

The output path is optional (defaults to the input filename with `.pdf`).
An existing output PDF is overwritten. The CLI returns a nonzero exit code on failure.

## Verification

Run `dotnet build ./PowerPointToPdf.slnx`. In the playground, run `npm run build` and
`npm run test:e2e` (installed Google Chrome required). UI tests mock conversion
responses. For real conversion, start the Compose services and use the curl
example above with a valid document. Fonts and LibreOffice layout may differ
from Microsoft Office; install matching fonts for closest layout.

To run the live browser/API check against Docker Compose:

```powershell
$env:POWERPOINT_TO_PDF_PROXY_URL = "http://localhost:8110"
$env:POWERPOINT_TO_PDF_LIVE_TEST = "1"
npm run test:e2e
```

Run these commands from the playground directory.

PDF output contains static slides; animations, transitions, audio and video are not preserved as interactive media. Hidden slides and speaker notes use LibreOffice export defaults.

The `powerpointtopdf.yml` workflow builds all projects and Docker images. It also installs LibreOffice on Linux, checks CLI output text, and runs browser tests against the real API/proxy. Release tags use `releases/powerpointtopdf/vMAJOR.MINOR.PATCH`.

Export uses the [LibreOffice PDF command-line interface](https://help.libreoffice.org/latest/en-US/text/shared/guide/pdf_params.html) with `impress_pdf_Export`.
