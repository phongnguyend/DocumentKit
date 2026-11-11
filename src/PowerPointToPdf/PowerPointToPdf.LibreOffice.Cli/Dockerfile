FROM mcr.microsoft.com/dotnet/sdk:11.0-resolute AS build
WORKDIR /src
COPY PowerPointToPdf.LibreOffice.Cli/PowerPointToPdf.LibreOffice.Cli.csproj PowerPointToPdf.LibreOffice.Cli/
RUN dotnet restore PowerPointToPdf.LibreOffice.Cli/PowerPointToPdf.LibreOffice.Cli.csproj
COPY PowerPointToPdf.LibreOffice.Cli/ PowerPointToPdf.LibreOffice.Cli/
COPY PowerPointToPdf.LibreOffice.Api/LibreOfficeConverter.cs PowerPointToPdf.LibreOffice.Api/
RUN dotnet publish PowerPointToPdf.LibreOffice.Cli/PowerPointToPdf.LibreOffice.Cli.csproj -c Release -o /app/publish --no-restore

FROM mcr.microsoft.com/dotnet/runtime:11.0-resolute AS final
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-impress fonts-dejavu-core fonts-liberation \
    && rm -rf /var/lib/apt/lists/*
ENV LIBREOFFICE_BIN=/usr/bin/soffice
WORKDIR /app
COPY --from=build /app/publish ./
ENV ASPNETCORE_HTTP_PORTS=8080
USER $APP_UID
EXPOSE 8080
ENTRYPOINT ["dotnet", "PowerPointToPdf.LibreOffice.Cli.dll"]
