FROM mcr.microsoft.com/dotnet/sdk:11.0-resolute AS build
WORKDIR /src

COPY HtmlToPdf.Playwright.Api.csproj ./
RUN dotnet restore

COPY . ./
RUN dotnet publish -c Release -o /app/publish --no-restore

FROM mcr.microsoft.com/dotnet/aspnet:11.0-resolute AS dotnet-runtime

# Keep this tag synchronized with the Microsoft.Playwright package version.
FROM mcr.microsoft.com/playwright/dotnet:v1.64.1-noble AS final
COPY --from=dotnet-runtime /usr/share/dotnet /usr/share/dotnet
ENV ASPNETCORE_HTTP_PORTS=8080

WORKDIR /app
COPY --from=build /app/publish ./

EXPOSE 8080
ENTRYPOINT ["dotnet", "HtmlToPdf.Playwright.Api.dll"]
