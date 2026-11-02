FROM mcr.microsoft.com/dotnet/sdk:10.0-noble AS build
WORKDIR /src
COPY PowerPointToPdf/PowerPointToPdf.PptxObjectKit.Api/ PowerPointToPdf/PowerPointToPdf.PptxObjectKit.Api/
COPY PptxObjectKit/Directory.Build.props PptxObjectKit/
COPY PptxObjectKit/PptxObjectKit/ PptxObjectKit/PptxObjectKit/
COPY PptxObjectKit/PptxObjectKit.Extensions.Pdf/ PptxObjectKit/PptxObjectKit.Extensions.Pdf/
RUN dotnet publish PowerPointToPdf/PowerPointToPdf.PptxObjectKit.Api -c Release -o /app/publish
FROM mcr.microsoft.com/dotnet/aspnet:10.0-noble
RUN apt-get update && apt-get install -y --no-install-recommends libfontconfig1 fonts-dejavu-core fonts-liberation && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY --from=build /app/publish ./
ENV ASPNETCORE_HTTP_PORTS=8080
USER $APP_UID
ENTRYPOINT ["dotnet", "PowerPointToPdf.PptxObjectKit.Api.dll"]
