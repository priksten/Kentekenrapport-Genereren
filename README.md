# Kentekenrapport-Genereren
Programma dat op basis van de invoer van een kenteken een bijbehorend ketekenrapport genereert

# Omschrijving:
De RDW heeft een API die de gebruiker in staat stelt om op basis van een kenteken van een auto gebruik te maken van data die over dat voertuig is opgeslagen.  Bij een gegeven kenteken zijn allerlei gegevens opgeslagen, denk hierbij aan merk, model, datum laatste tenaanstelling, vervaldatum APK, kentekengewicht e.d. Het doel van dit project is om een programma te schrijven die de gebruiker in staat stelt om op basis van een kenteken zoveel mogelijk informatie uit het RDW-systeem te halen. Hierbij is het de bedoeling dat deze gegevens op basis van meerdere filtermogelijkheden kunnen worden weergegeven. Hierbij moet er een nette userinterface ontworpen worden en de mogelijkheid om een kentekenrapport in MS-Word en/of PDF te maken.

# Ontwikkelingsfasen:
De ontwikkeling van dit programma is volgende de volgende fasen verlopen:

Fase 1: Programma in Terminal 
Ontwikkeling programma voor het opvragen van autogegevens op basis van het kenteken via de RDW API. Dit programa werkt in de terminal. In de eindversie is het mogelijk om gegevens te filteren op basis van de volgende vijf categorieën: 
1. Basisgegevens
2. Registratie 
3. Motor 
4. Milieu
5. Maten en Gewichten
Daarnaast is er de mogelijkheid om van de gewenste gegevens een rapport te krijgen, zowel in Word als in PDF. Zie screenshots voor hoe dit programma werkt.

Fase 2: Ontwikkeling Userinterface
In deze fase wordt er een userinterface ontwikkeld voor het programma dat in fase 1 is ontwikkeld. Dit programma biedt dezelfde functionaliteiten als het programma uit fase 1. Het verschil is dat de terminal in dit programma niet meer gebruikt wordt. Zie ook screenshots van hoe dit programma eruitziet en werkt. 
Tijdens het ontwikkelingsproces van de Userinterface heb ik de code van het eindresultaat van fase 1 ingrijpend aangepast. Het verschil tussen het eindproduct van fase 1 en het eindproduct van fase 2 is dat ik in het eindproduct van fase 2 heb geprobeerd de code aanzienlijk compacter te krijgen door het gebruik van functies en modules. 

# Programmeertaal + versie: 
Python

# Gebruikte modules:
Docx

Docx2pdf

Pandas

RDW

Tkinter

# Screenshots:

Versie zonder GUI:
![image](https://github.com/priksten/Kentekenrapport-Genereren/assets/85739742/537a10d9-3ab8-46af-86ee-676b7b0940b4)
![image](https://github.com/priksten/Kentekenrapport-Genereren/assets/85739742/e9372f83-f790-4678-b575-ec6ae230a4e5)
![image](https://github.com/priksten/Kentekenrapport-Genereren/assets/85739742/5494780a-0c96-4189-8324-0e3ca86ac8ff)
![image](https://github.com/priksten/Kentekenrapport-Genereren/assets/85739742/9523d1d6-9c97-4111-b92a-6efecea4089a)

Wanneer de gebruiker een kenteken heeft ingevoerd, wordt er altijd een kentekenrapport in MS-Word gegenereerd. Het programma vraagt de gebruiker of hij het kentekenrapport ook als PDF wil. Zo ja, dan converteert het programma het Word-bestand naar PDF. De bestandsnaam is in beide gevallen het door de gebruiker ingevoerde kenteken. 
![image](https://github.com/priksten/Kentekenrapport-Genereren/assets/85739742/0a0ded8a-0368-436b-aac5-fe31b0edd8ed)
 
Versie met GUI:
De algemene lay-out van de grafische gebruikersinterface:
![image](https://github.com/priksten/Kentekenrapport-Genereren/assets/85739742/ac8cca87-a18c-4872-8add-c5e4d4621439)

Linksboven kan de gebruiker een kenteken invoeren. Daaronder is er een menu waarmee de gebruiker kan kiezen welke categorieën informatie er in het scherm worden weergegeven. Daaronder zit een knop waarmee de gebruiker een kentekenrapport in zowel MS-Word als PDF kan laten genereren. 

Screenshots met alle categorieën:
![image](https://github.com/priksten/Kentekenrapport-Genereren/assets/85739742/589f766c-9c8a-48fd-a28f-bbd217046bd9)
 
Screenshot met daarin de weergave van de informatie uit een aantal geselecteede categorieën:
![image](https://github.com/priksten/Kentekenrapport-Genereren/assets/85739742/5b263e7c-2f1c-4a2b-a028-2dd1f90a7843)
