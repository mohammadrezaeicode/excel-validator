FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

COPY validator.csproj .
RUN dotnet restore

COPY . .
RUN dotnet publish -c Release -o /out

FROM mcr.microsoft.com/dotnet/aspnet:8.0

RUN apt-get update && apt-get install -y \
    libreoffice \
    libfontconfig1 \
    libfreetype6 \
    fonts-liberation \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY --from=build /out .

EXPOSE 8080
ENV ASPNETCORE_URLS=http://+:8080

ENTRYPOINT ["dotnet", "validator.dll"]
