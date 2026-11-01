FROM mcr.microsoft.com/dotnet/sdk:10.0-noble AS build
WORKDIR /src
COPY PowerPointToPdf.Proxy/PowerPointToPdf.Proxy.csproj PowerPointToPdf.Proxy/
RUN dotnet restore PowerPointToPdf.Proxy/PowerPointToPdf.Proxy.csproj
COPY PowerPointToPdf.Proxy/ PowerPointToPdf.Proxy/
RUN dotnet publish PowerPointToPdf.Proxy/PowerPointToPdf.Proxy.csproj -c Release -o /app/publish --no-restore

FROM mcr.microsoft.com/dotnet/aspnet:10.0-noble AS final
WORKDIR /app
COPY --from=build /app/publish ./
ENV ASPNETCORE_HTTP_PORTS=8080
USER $APP_UID
EXPOSE 8080
ENTRYPOINT ["dotnet", "PowerPointToPdf.Proxy.dll"]
