FROM mcr.microsoft.com/dotnet/sdk:10.0-noble AS build
WORKDIR /src
COPY WordToPdf/WordToPdf.DocxObjectKit.Api/ WordToPdf/WordToPdf.DocxObjectKit.Api/
COPY DocxObjectKit/Directory.Build.props DocxObjectKit/
COPY DocxObjectKit/DocxObjectKit/ DocxObjectKit/DocxObjectKit/
COPY DocxObjectKit/DocxObjectKit.Extensions.Pdf/ DocxObjectKit/DocxObjectKit.Extensions.Pdf/
RUN dotnet publish WordToPdf/WordToPdf.DocxObjectKit.Api -c Release -o /app/publish
FROM mcr.microsoft.com/dotnet/aspnet:10.0-noble
RUN apt-get update && apt-get install -y --no-install-recommends libfontconfig1 fonts-dejavu-core fonts-liberation fonts-crosextra-carlito fonts-crosextra-caladea && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY --from=build /app/publish ./
ENV ASPNETCORE_HTTP_PORTS=8080
USER $APP_UID
ENTRYPOINT ["dotnet", "WordToPdf.DocxObjectKit.Api.dll"]
