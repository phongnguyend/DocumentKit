FROM mcr.microsoft.com/dotnet/sdk:10.0-noble AS build
WORKDIR /src
COPY PowerPointToPdf.LibreOffice.Api/PowerPointToPdf.LibreOffice.Api.csproj PowerPointToPdf.LibreOffice.Api/
RUN dotnet restore PowerPointToPdf.LibreOffice.Api/PowerPointToPdf.LibreOffice.Api.csproj
COPY PowerPointToPdf.LibreOffice.Api/ PowerPointToPdf.LibreOffice.Api/
RUN dotnet publish PowerPointToPdf.LibreOffice.Api/PowerPointToPdf.LibreOffice.Api.csproj -c Release -o /app/publish --no-restore

FROM mcr.microsoft.com/dotnet/aspnet:10.0-noble AS final
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-impress fonts-dejavu-core fonts-liberation \
    && rm -rf /var/lib/apt/lists/*
ENV LIBREOFFICE_BIN=/usr/bin/soffice
WORKDIR /app
COPY --from=build /app/publish ./
ENV ASPNETCORE_HTTP_PORTS=8080
USER $APP_UID
EXPOSE 8080
ENTRYPOINT ["dotnet", "PowerPointToPdf.LibreOffice.Api.dll"]
