# PPTXcreator
[![CodeFactor Grade](https://img.shields.io/codefactor/grade/github/Sionkerk-Techniek/PPTXcreator)](https://www.codefactor.io/repository/github/sionkerk-techniek/pptxcreator) [![Build workflow](https://img.shields.io/github/workflow/status/Sionkerk-Techniek/PPTXcreator/.NET%20Framework%20Build)](https://github.com/Sionkerk-Techniek/PPTXcreator/actions/workflows/netframeworkbuild.yaml?query=branch%3Amaster) [![Supported platforms](https://img.shields.io/badge/platform-windows-informational)](#) [![Latest version](https://img.shields.io/github/v/tag/Sionkerk-Techniek/PPTXcreator?label=laatste%20versie)](https://github.com/Sionkerk-Techniek/PPTXcreator/releases/latest)

## Gebruik
Start het programma door `PPTXcreator.exe` uit te voeren. Mogelijk verschijnt er een waarschuwingsscherm van Windows, klik dan op 'Meer informatie' en 'Toch uitvoeren'. Controleer bij de eerste keer uitvoeren eerst de instellingen, selecteer daar de goede bestanden met `...`. Vervolgens kun je de informatie van de dienst invoeren in de andere drie tabs.

Voor het automatisch invullen van velden is er een `.json` bestand met diensteninformatie nodig die van de Google Drive kan worden gedownload, en de presentatietemplates zijn daar ook te vinden. Als je de `.json` zelf wil maken kun je [Sionkerk-Techniek/json-constructor](https://github.com/Sionkerk-Techniek/json-constructor) gebruiken.

## Installatie
Na het downloaden moet het zip-bestand worden uitgepakt in een willekeurige folder. Als je de bestanden zelf wil maken met de broncode kan dat met `msbuild PPTXcreator.sln /t:Build;Package /p:Configuration=Release`.
