FROM mcr.microsoft.com/powershell:lts-7.4-ubuntu-22.04

RUN pwsh -c "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted;  \
    Install-Module -Name ExchangeOnlineManagement -Force"

WORKDIR /app
EXPOSE 8080
COPY ExchangeGroupManager.ps1 .
CMD ["pwsh", "-File", "./ExchangeGroupManager.ps1"]
