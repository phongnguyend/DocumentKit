FROM mcr.microsoft.com/dotnet/sdk:11.0-resolute AS build
WORKDIR /src
COPY . .
RUN dotnet publish PptxObjectKit.Playground.Api/PptxObjectKit.Playground.Api.csproj -c Release -o /app/publish
FROM mcr.microsoft.com/dotnet/aspnet:11.0-resolute
WORKDIR /app
COPY --from=build /app/publish ./
ENV ASPNETCORE_HTTP_PORTS=8080
USER $APP_UID
ENTRYPOINT ["dotnet", "PptxObjectKit.Playground.Api.dll"]
