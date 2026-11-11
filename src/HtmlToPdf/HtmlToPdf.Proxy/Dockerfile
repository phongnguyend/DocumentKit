FROM mcr.microsoft.com/dotnet/sdk:11.0-resolute AS build
WORKDIR /src

COPY HtmlToPdf.Proxy.csproj ./
RUN dotnet restore

COPY . ./
RUN dotnet publish -c Release -o /app/publish --no-restore

FROM mcr.microsoft.com/dotnet/aspnet:11.0-resolute AS final
ENV ASPNETCORE_HTTP_PORTS=8080

WORKDIR /app
COPY --from=build /app/publish ./

EXPOSE 8080
ENTRYPOINT ["dotnet", "HtmlToPdf.Proxy.dll"]
