# DocumentKit

Document libraries, conversion services, CLIs, and browser playgrounds built with
.NET 10 and React. Each solution has its own build workflow and setup instructions.

- [XlsxObjectKit](src/XlsxObjectKit/README.md): inspect and edit XLSX workbooks in a native worksheet grid, manage worksheet tabs, format cells, and save changes through a .NET API or CLI.
- [DocxObjectKit](src/DocxObjectKit/README.md): inspect, edit and preview Word documents, manage text and images, extract Markdown, and export PDFs with a native .NET engine.
- [PptxObjectKit](src/PptxObjectKit/README.md): inspect, edit and preview PowerPoint presentations, manage slides, text and images, extract Markdown, and export PDFs with a native .NET engine.
- [PdfObjectKit](src/PdfObjectKit/README.md): inspect, edit and render PDFs; manage pages, text, images, forms and annotations; extract Markdown; and export editable Word documents and PowerPoint slides.
- [HtmlToPdf](src/HtmlToPdf/README.md): convert HTML files and webpage URLs to PDF using Chrome, Playwright, or PuppeteerSharp.
- [ExcelToPdf](src/ExcelToPdf/README.md): convert Excel workbooks to PDF using LibreOffice.
- [WordToPdf](src/WordToPdf/README.md): convert Word documents to PDF using LibreOffice or the native DocxObjectKit engine.
- [PowerPointToPdf](src/PowerPointToPdf/README.md): convert PowerPoint presentations to PDF using LibreOffice or the native PptxObjectKit engine.

All eight playgrounds offer the same **Light / Dark / System** switcher in the
header. Preferences persist in browser storage and synchronize across tabs on the
same origin; separate ports or hosts keep separate preferences. Document previews
and worksheet cell formatting retain their original colors. Run `npm run test:theme`
inside a playground directory to check theme behavior without starting its backend.

## Build workflows

GitHub Actions workflows run on pushes and pull requests matching their configured
paths, and can also be started manually from the Actions tab. All use .NET 10;
browser-test jobs also use Node.js 22 and Playwright.

| Workflow | Build and validation | Release publishing |
| --- | --- | --- |
| [XlsxObjectKit](.github/workflows/xlsxobjectkit.yml) | Release solution build; package/edit and Open XML validation tests; UI build and browser tests; Compose image builds | Not configured |
| [DocxObjectKit](.github/workflows/docxobjectkit.yml) | Release solution build; package/edit and native PDF tests; UI build and browser tests | Not configured |
| [PptxObjectKit](.github/workflows/pptxobjectkit.yml) | Release solution build; core and native PDF tests; UI build and browser tests; Compose image builds | Not configured |
| [PdfObjectKit](.github/workflows/pdfobjectkit.yml) | Core, Word and PowerPoint extension, Markdown snapshot, CLI and UI-helper tests; UI build and browser tests; Docker image builds | NuGet packages and Docker images |
| [HtmlToPdf](.github/workflows/htmltopdf.yml) | Release solution build; Docker image builds; URL and HTML conversion integration tests for the engine APIs | Docker images |
| [ExcelToPdf](.github/workflows/exceltopdf.yml) | Release solution build and Docker image builds | Docker images |
| [WordToPdf](.github/workflows/wordtopdf.yml) | Release solution build and Docker image builds, including the native DocxObjectKit API | Docker images |
| [PowerPointToPdf](.github/workflows/powerpointtopdf.yml) | Release solution build; CLI and API conversion checks; UI build and browser tests; Docker image builds, including the native PptxObjectKit API | Docker images |

WordToPdf also watches changes to its DocxObjectKit engine dependencies, and
PowerPointToPdf watches the corresponding PptxObjectKit dependencies. The workflow
files define the exact path filters, test commands, and uploaded diagnostics.

### Release Docker images

HtmlToPdf, ExcelToPdf, WordToPdf, PowerPointToPdf, and PdfObjectKit build their
workflow image matrices for Linux AMD64. Each image has a separate build-cache
scope. Branch and pull-request builds use local `:ci` tags; release-tag runs
publish the built images to Docker Hub. DocxObjectKit and PptxObjectKit currently
provide local Compose builds without standalone Docker Hub release publishing.

PdfObjectKit also publishes its core, Markdown, Word, and PowerPoint extension NuGet packages on
release tags, using the `NUGET_API_KEY` Actions secret. See
[PdfObjectKit Docker Compose](src/PdfObjectKit/README.md#docker-compose) for local
builds, published deployments, and CLI usage.

Configure these repository **Actions secrets**:

- `DOCKERHUB_USERNAME`: the Docker Hub account used to sign in.
- `DOCKERHUB_TOKEN`: a Docker Hub access token with write access to the target repositories.

Optionally set the **Actions variable** `DOCKERHUB_NAMESPACE` to a Docker Hub
organization. It defaults to `DOCKERHUB_USERNAME`; the account must have write
access to repositories in that namespace.

Following the release-tag convention in `pdfobjectkit.yml`, push one of:

| Solution | Git tag example | Published image example |
| --- | --- | --- |
| HtmlToPdf | `releases/htmltopdf/v1.2.3` | `<namespace>/html-to-pdf-playwright-api:1.2.3` |
| ExcelToPdf | `releases/exceltopdf/v1.2.3` | `<namespace>/excel-to-pdf-libreoffice-api:1.2.3` |
| WordToPdf | `releases/wordtopdf/v1.2.3` | `<namespace>/word-to-pdf-libreoffice-api:1.2.3` |
| PowerPointToPdf | `releases/powerpointtopdf/v1.2.3` | `<namespace>/powerpoint-to-pdf-libreoffice-api:1.2.3` |
| PdfObjectKit | `releases/pdfobjectkit/v1.2.3` | `<namespace>/pdfobjectkit-playground-api:1.2.3` |

```powershell
git tag releases/htmltopdf/v1.2.3
git push origin releases/htmltopdf/v1.2.3
```

Each release publishes every image in that solution's workflow matrix, including
engine APIs, CLIs, proxy, and playground. Repository names replace the slash in
the matrix image name with a hyphen: `html-to-pdf/proxy` becomes
`<namespace>/html-to-pdf-proxy:1.2.3`. Images target Linux AMD64.
Prereleases such as `v1.2.3-rc.1` publish the `1.2.3-rc.1` image tag.
Stable releases also publish `latest` from the same built image as the version
tag, for every engine API, CLI, proxy, and playground. Prereleases do not update
`latest`. Rerunning a stable release republishes both tags, so rerun the newest
stable release when backfilling `latest` for existing images.
Use a new version tag for each release. A failed matrix job can be rerun to
finish publishing that version. Manual dispatch on a release tag also publishes;
manual dispatch on a branch only builds.

For local .NET and service-image builds, run these commands from the relevant
solution folder (replace `HtmlToPdf` with the solution name):

```powershell
dotnet build HtmlToPdf.slnx --configuration Release
docker compose build
```

Compose builds the services declared by each solution. The release workflow
matrices additionally include CLI images where available.

For local object-kit playgrounds, run `docker compose up --build` from
`src/DocxObjectKit`, `src/PptxObjectKit`, or `src/PdfObjectKit`:

| Playground | Local URL | Instructions |
| --- | --- | --- |
| DocxObjectKit | http://localhost:8126 | [Setup and native layout capabilities](src/DocxObjectKit/README.md) |
| PptxObjectKit | http://localhost:8124 | [Setup and presentation features](src/PptxObjectKit/README.md) |
| PdfObjectKit | http://localhost:8074 | [Setup and PDF features](src/PdfObjectKit/README.md) |

### Run published Docker images

The four conversion-service solutions and PdfObjectKit include a
`compose.published.yaml` that pulls their services from Docker Hub without local
builds. For example:

```powershell
cd src/ExcelToPdf # Or src/HtmlToPdf or src/WordToPdf or src/PowerPointToPdf from the repository root
Copy-Item .env.sample .env # Only if .env does not already exist
```

Edit `.env` in that solution folder:

```dotenv
DOCKERHUB_NAMESPACE=your-dockerhub-account-or-organization
IMAGE_TAG=latest
```

Compose loads this file automatically. Local `.env` files are ignored by Git;
`.env.sample` provides the shared template. Then start the services:

```powershell
docker compose -f compose.published.yaml up -d
```

`DOCKERHUB_NAMESPACE` must match the namespace used by the release workflow.
Images default to `latest` and are pulled on every `up` to pick up republished
tags. To pin a release, set `IMAGE_TAG=1.2.3` in `.env` before running the command;
use the image version without the Git tag prefix or `v`.

The published playgrounds are available at http://localhost:8074 (PdfObjectKit),
http://localhost:8084 (HTML),
http://localhost:8094 (Excel), http://localhost:8104 (Word), and
http://localhost:8114 (PowerPoint).
Published Compose uses the same project names and ports as local Compose, so
it replaces that solution's existing containers when switching deployment modes.
Images require Linux containers and target AMD64.

To stop the published deployment, run
`docker compose -f compose.published.yaml down` from its solution folder.
