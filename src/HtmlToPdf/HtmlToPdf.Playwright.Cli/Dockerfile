FROM mcr.microsoft.com/dotnet/sdk:11.0-resolute AS build
WORKDIR /src

COPY HtmlToPdf.Playwright.Cli.csproj ./
RUN dotnet restore

COPY . ./
RUN dotnet publish -c Release -o /app/publish --no-restore

FROM mcr.microsoft.com/dotnet/runtime:11.0-resolute AS dotnet-runtime

# Keep this tag synchronized with the Microsoft.Playwright package version.
FROM mcr.microsoft.com/playwright/dotnet:v1.64.1-noble AS final
COPY --from=dotnet-runtime /usr/share/dotnet /usr/share/dotnet

WORKDIR /app
COPY --from=build /app/publish ./

WORKDIR /output
ENTRYPOINT ["dotnet", "/app/HtmlToPdf.Playwright.Cli.dll"]
