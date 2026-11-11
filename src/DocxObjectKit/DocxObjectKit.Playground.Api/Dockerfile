FROM mcr.microsoft.com/dotnet/sdk:11.0-resolute AS build
WORKDIR /src
COPY . .
RUN dotnet publish DocxObjectKit.Playground.Api -c Release -o /app/publish
FROM mcr.microsoft.com/dotnet/aspnet:11.0-resolute
RUN apt-get update && apt-get install -y --no-install-recommends libfontconfig1 fonts-dejavu-core fonts-liberation fonts-crosextra-carlito fonts-crosextra-caladea && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY --from=build /app/publish ./
ENV ASPNETCORE_HTTP_PORTS=8080
USER $APP_UID
ENTRYPOINT ["dotnet", "DocxObjectKit.Playground.Api.dll"]
