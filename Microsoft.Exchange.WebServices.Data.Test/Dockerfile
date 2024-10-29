FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
#USER $APP_UID // Пришлось установить пакет под root
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
ARG BUILD_CONFIGURATION=Release
WORKDIR /src
COPY ["Microsoft.Exchange.WebServices.Data.Test/Microsoft.Exchange.WebServices.Data.Test.csproj", "Microsoft.Exchange.WebServices.Data.Test/"]
COPY ["Microsoft.Exchange.WebServices.Data/Microsoft.Exchange.WebServices.Data.csproj", "Microsoft.Exchange.WebServices.Data/"]
RUN dotnet restore "Microsoft.Exchange.WebServices.Data.Test/Microsoft.Exchange.WebServices.Data.Test.csproj"
COPY . .
WORKDIR "/src/Microsoft.Exchange.WebServices.Data.Test"
RUN dotnet build "Microsoft.Exchange.WebServices.Data.Test.csproj" -c $BUILD_CONFIGURATION -o /app/build

FROM build AS publish
ARG BUILD_CONFIGURATION=Release
RUN dotnet publish "Microsoft.Exchange.WebServices.Data.Test.csproj" -c $BUILD_CONFIGURATION -o /app/publish /p:UseAppHost=false

FROM base AS final

# Библиотека нужна для работы NTLM авторизации
RUN apt-get update && apt-get install -y gss-ntlmssp

WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Microsoft.Exchange.WebServices.Data.Test.dll"]
