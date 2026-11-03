FROM mcr.microsoft.com/dotnet/sdk:10.0-noble AS build
WORKDIR /src
COPY . .
RUN dotnet publish XlsxObjectKit.Cli -c Release -o /app/publish
FROM mcr.microsoft.com/dotnet/runtime:10.0-noble
WORKDIR /app
COPY --from=build /app/publish ./
ENV ASPNETCORE_HTTP_PORTS=8080
USER $APP_UID
ENTRYPOINT ["dotnet", "XlsxObjectKit.Cli.dll"]
