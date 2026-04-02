# ============================================================
# Script : Creation du Global OptionSet - Departements de France
# Version : 4.0 - SDK PowerShell - Creation puis ajout un par un
# Prerequis : Module Microsoft.Xrm.Data.PowerShell
# ============================================================

# --- INSTALLATION DU MODULE (une seule fois) ---
# Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser -Force

Clear-Host
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  CREATION DU GLOBAL OPTIONSET - DEPARTEMENTS DE FRANCE" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ============================================================
# CHARGEMENT DU MODULE
# ============================================================
Write-Host "Chargement du module..." -ForegroundColor Gray
Import-Module Microsoft.Xrm.Data.PowerShell -ErrorAction Stop
Write-Host "Module charge." -ForegroundColor Gray
Write-Host ""

# ============================================================
# ETAPE 1 : CONNEXION A L'ENVIRONNEMENT
# ============================================================
Write-Host "ETAPE 1 : Connexion a Dataverse" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray
Write-Host ""
Write-Host "Exemple : https://monorg.crm4.dynamics.com" -ForegroundColor Gray
Write-Host ""
$envUrl = Read-Host "URL de ton environnement Dataverse"

if ([string]::IsNullOrWhiteSpace($envUrl)) {
    Write-Host "URL vide. Arret du script." -ForegroundColor Red
    exit
}

$envUrl = $envUrl.TrimEnd('/')

Write-Host ""
Write-Host "Connexion a $envUrl..." -ForegroundColor White
Write-Host ""

$conn = Connect-CrmOnline -ServerUrl $envUrl -ForceOAuth

if (-not $conn -or -not $conn.IsReady) {
    Write-Host "Echec de connexion. Arret du script." -ForegroundColor Red
    exit
}

Write-Host ""
Write-Host "Connecte a : $($conn.ConnectedOrgFriendlyName)" -ForegroundColor Green
Write-Host ""

# ============================================================
# ETAPE 2 : SELECTION DE LA SOLUTION
# ============================================================
Write-Host "ETAPE 2 : Selection de la solution" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray
Write-Host ""
Write-Host "Recuperation des solutions non managees..." -ForegroundColor White

$fetchXml = @"
<fetch>
  <entity name="solution">
    <attribute name="solutionid" />
    <attribute name="friendlyname" />
    <attribute name="uniquename" />
    <attribute name="version" />
    <attribute name="publisherid" />
    <filter>
      <condition attribute="ismanaged" operator="eq" value="0" />
      <condition attribute="isvisible" operator="eq" value="1" />
    </filter>
    <link-entity name="publisher" from="publisherid" to="publisherid" alias="pub">
      <attribute name="customizationprefix" />
      <attribute name="friendlyname" />
    </link-entity>
    <order attribute="friendlyname" />
  </entity>
</fetch>
"@

$solutions = Get-CrmRecordsByFetch -conn $conn -Fetch $fetchXml

if ($solutions.Count -eq 0) {
    Write-Host "Aucune solution non managee trouvee. Arret du script." -ForegroundColor Red
    exit
}

Write-Host ""
Write-Host "Solutions non managees disponibles :" -ForegroundColor Cyan
Write-Host ""

$index = 1
$solutionList = @()

foreach ($sol in $solutions.CrmRecords) {
    $prefix = $sol."pub.customizationprefix"
    $publisherName = $sol."pub.friendlyname"
    $displayName = $sol.friendlyname
    $uniqueName = $sol.uniquename
    $version = $sol.version
    
    $solutionList += @{
        Index = $index
        FriendlyName = $displayName
        UniqueName = $uniqueName
        Prefix = $prefix
        PublisherName = $publisherName
        Version = $version
        SolutionId = $sol.solutionid
    }
    
    Write-Host "  [$index] $displayName" -ForegroundColor White
    Write-Host "      Prefixe: $prefix | Editeur: $publisherName" -ForegroundColor Gray
    Write-Host ""
    
    $index++
}

Write-Host "------------------------------------------------------------" -ForegroundColor Gray
$choix = Read-Host "Entre le numero de la solution (1-$($solutionList.Count))"

$choixInt = 0
if (-not [int]::TryParse($choix, [ref]$choixInt) -or $choixInt -lt 1 -or $choixInt -gt $solutionList.Count) {
    Write-Host "Choix invalide. Arret du script." -ForegroundColor Red
    exit
}

$selectedSolution = $solutionList | Where-Object { $_.Index -eq $choixInt }
$prefix = $selectedSolution.Prefix
$solutionUniqueName = $selectedSolution.UniqueName

Write-Host ""
Write-Host "Solution selectionnee : $($selectedSolution.FriendlyName)" -ForegroundColor Green
Write-Host "Prefixe : $prefix" -ForegroundColor Green
Write-Host ""

# ============================================================
# ETAPE 3 : NOM DE L'OPTIONSET
# ============================================================
Write-Host "ETAPE 3 : Nom de l'OptionSet" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray
Write-Host ""
Write-Host "Le prefixe '$prefix' sera ajoute automatiquement." -ForegroundColor Gray
Write-Host "Exemple : si tu tapes 'departement', le nom sera '${prefix}_departement'" -ForegroundColor Gray
Write-Host ""

$optionSetExists = $true

while ($optionSetExists) {
    $optionSetSuffix = Read-Host "Nom de l'OptionSet (sans prefixe)"
    
    if ([string]::IsNullOrWhiteSpace($optionSetSuffix)) {
        Write-Host "Nom vide. Arret du script." -ForegroundColor Red
        exit
    }
    
    # Nettoyer le nom (minuscules, underscores)
    $optionSetSuffix = $optionSetSuffix.ToLower().Trim() -replace '\s+', '_'
    $optionSetName = "${prefix}_${optionSetSuffix}"
    $optionSetDisplayName = $optionSetSuffix.Substring(0,1).ToUpper() + $optionSetSuffix.Substring(1)
    
    Write-Host ""
    Write-Host "Verification si '$optionSetName' existe deja..." -ForegroundColor White
    
    # Verifier si l'OptionSet existe
    try {
        $retrieveRequest = New-Object Microsoft.Xrm.Sdk.Messages.RetrieveOptionSetRequest
        $retrieveRequest.Name = $optionSetName
        $response = $conn.Execute($retrieveRequest)
        
        # Si on arrive ici, l'OptionSet existe
        Write-Host "L'OptionSet '$optionSetName' existe deja !" -ForegroundColor Red
        Write-Host "Choisis un autre nom." -ForegroundColor Yellow
        Write-Host ""
        $optionSetExists = $true
    }
    catch {
        # L'OptionSet n'existe pas, on peut continuer
        Write-Host "OK, '$optionSetName' est disponible." -ForegroundColor Green
        $optionSetExists = $false
    }
}

Write-Host ""
Write-Host "   Nom technique : $optionSetName" -ForegroundColor White
Write-Host "   Nom affiche   : $optionSetDisplayName" -ForegroundColor White
Write-Host ""

$confirm = Read-Host "Confirmer la creation ? (O/N)"
if ($confirm -notin @("O", "o", "Y", "y")) {
    Write-Host "Annule." -ForegroundColor Yellow
    exit
}

Write-Host ""

# ============================================================
# LISTE DES DEPARTEMENTS (Value = numero du departement)
# ============================================================
$departements = @(
    @{ Value = 1;   Label = "01 - Ain" }
    @{ Value = 2;   Label = "02 - Aisne" }
    @{ Value = 3;   Label = "03 - Allier" }
    @{ Value = 4;   Label = "04 - Alpes-de-Haute-Provence" }
    @{ Value = 5;   Label = "05 - Hautes-Alpes" }
    @{ Value = 6;   Label = "06 - Alpes-Maritimes" }
    @{ Value = 7;   Label = "07 - Ardeche" }
    @{ Value = 8;   Label = "08 - Ardennes" }
    @{ Value = 9;   Label = "09 - Ariege" }
    @{ Value = 10;  Label = "10 - Aube" }
    @{ Value = 11;  Label = "11 - Aude" }
    @{ Value = 12;  Label = "12 - Aveyron" }
    @{ Value = 13;  Label = "13 - Bouches-du-Rhone" }
    @{ Value = 14;  Label = "14 - Calvados" }
    @{ Value = 15;  Label = "15 - Cantal" }
    @{ Value = 16;  Label = "16 - Charente" }
    @{ Value = 17;  Label = "17 - Charente-Maritime" }
    @{ Value = 18;  Label = "18 - Cher" }
    @{ Value = 19;  Label = "19 - Correze" }
    @{ Value = 201; Label = "2A - Corse-du-Sud" }
    @{ Value = 202; Label = "2B - Haute-Corse" }
    @{ Value = 21;  Label = "21 - Cote-d-Or" }
    @{ Value = 22;  Label = "22 - Cotes-d-Armor" }
    @{ Value = 23;  Label = "23 - Creuse" }
    @{ Value = 24;  Label = "24 - Dordogne" }
    @{ Value = 25;  Label = "25 - Doubs" }
    @{ Value = 26;  Label = "26 - Drome" }
    @{ Value = 27;  Label = "27 - Eure" }
    @{ Value = 28;  Label = "28 - Eure-et-Loir" }
    @{ Value = 29;  Label = "29 - Finistere" }
    @{ Value = 30;  Label = "30 - Gard" }
    @{ Value = 31;  Label = "31 - Haute-Garonne" }
    @{ Value = 32;  Label = "32 - Gers" }
    @{ Value = 33;  Label = "33 - Gironde" }
    @{ Value = 34;  Label = "34 - Herault" }
    @{ Value = 35;  Label = "35 - Ille-et-Vilaine" }
    @{ Value = 36;  Label = "36 - Indre" }
    @{ Value = 37;  Label = "37 - Indre-et-Loire" }
    @{ Value = 38;  Label = "38 - Isere" }
    @{ Value = 39;  Label = "39 - Jura" }
    @{ Value = 40;  Label = "40 - Landes" }
    @{ Value = 41;  Label = "41 - Loir-et-Cher" }
    @{ Value = 42;  Label = "42 - Loire" }
    @{ Value = 43;  Label = "43 - Haute-Loire" }
    @{ Value = 44;  Label = "44 - Loire-Atlantique" }
    @{ Value = 45;  Label = "45 - Loiret" }
    @{ Value = 46;  Label = "46 - Lot" }
    @{ Value = 47;  Label = "47 - Lot-et-Garonne" }
    @{ Value = 48;  Label = "48 - Lozere" }
    @{ Value = 49;  Label = "49 - Maine-et-Loire" }
    @{ Value = 50;  Label = "50 - Manche" }
    @{ Value = 51;  Label = "51 - Marne" }
    @{ Value = 52;  Label = "52 - Haute-Marne" }
    @{ Value = 53;  Label = "53 - Mayenne" }
    @{ Value = 54;  Label = "54 - Meurthe-et-Moselle" }
    @{ Value = 55;  Label = "55 - Meuse" }
    @{ Value = 56;  Label = "56 - Morbihan" }
    @{ Value = 57;  Label = "57 - Moselle" }
    @{ Value = 58;  Label = "58 - Nievre" }
    @{ Value = 59;  Label = "59 - Nord" }
    @{ Value = 60;  Label = "60 - Oise" }
    @{ Value = 61;  Label = "61 - Orne" }
    @{ Value = 62;  Label = "62 - Pas-de-Calais" }
    @{ Value = 63;  Label = "63 - Puy-de-Dome" }
    @{ Value = 64;  Label = "64 - Pyrenees-Atlantiques" }
    @{ Value = 65;  Label = "65 - Hautes-Pyrenees" }
    @{ Value = 66;  Label = "66 - Pyrenees-Orientales" }
    @{ Value = 67;  Label = "67 - Bas-Rhin" }
    @{ Value = 68;  Label = "68 - Haut-Rhin" }
    @{ Value = 69;  Label = "69 - Rhone" }
    @{ Value = 70;  Label = "70 - Haute-Saone" }
    @{ Value = 71;  Label = "71 - Saone-et-Loire" }
    @{ Value = 72;  Label = "72 - Sarthe" }
    @{ Value = 73;  Label = "73 - Savoie" }
    @{ Value = 74;  Label = "74 - Haute-Savoie" }
    @{ Value = 75;  Label = "75 - Paris" }
    @{ Value = 76;  Label = "76 - Seine-Maritime" }
    @{ Value = 77;  Label = "77 - Seine-et-Marne" }
    @{ Value = 78;  Label = "78 - Yvelines" }
    @{ Value = 79;  Label = "79 - Deux-Sevres" }
    @{ Value = 80;  Label = "80 - Somme" }
    @{ Value = 81;  Label = "81 - Tarn" }
    @{ Value = 82;  Label = "82 - Tarn-et-Garonne" }
    @{ Value = 83;  Label = "83 - Var" }
    @{ Value = 84;  Label = "84 - Vaucluse" }
    @{ Value = 85;  Label = "85 - Vendee" }
    @{ Value = 86;  Label = "86 - Vienne" }
    @{ Value = 87;  Label = "87 - Haute-Vienne" }
    @{ Value = 88;  Label = "88 - Vosges" }
    @{ Value = 89;  Label = "89 - Yonne" }
    @{ Value = 90;  Label = "90 - Territoire de Belfort" }
    @{ Value = 91;  Label = "91 - Essonne" }
    @{ Value = 92;  Label = "92 - Hauts-de-Seine" }
    @{ Value = 93;  Label = "93 - Seine-Saint-Denis" }
    @{ Value = 94;  Label = "94 - Val-de-Marne" }
    @{ Value = 95;  Label = "95 - Val-d-Oise" }
    @{ Value = 971; Label = "971 - Guadeloupe" }
    @{ Value = 972; Label = "972 - Martinique" }
    @{ Value = 973; Label = "973 - Guyane" }
    @{ Value = 974; Label = "974 - La Reunion" }
    @{ Value = 976; Label = "976 - Mayotte" }
)

# ============================================================
# ETAPE 5 : CREATION DU GLOBAL OPTIONSET VIDE
# ============================================================
Write-Host "ETAPE 5 : Creation du Global OptionSet..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray
Write-Host ""

# Creer l'OptionSet vide avec une seule option temporaire
$firstOption = New-Object Microsoft.Xrm.Sdk.Metadata.OptionMetadata
$firstOption.Value = $departements[0].Value
$firstOption.Label = New-Object Microsoft.Xrm.Sdk.Label($departements[0].Label, 1036)

$optionsList = New-Object Microsoft.Xrm.Sdk.Metadata.OptionMetadataCollection
$optionsList.Add($firstOption)

$optionSetMetadata = New-Object Microsoft.Xrm.Sdk.Metadata.OptionSetMetadata
$optionSetMetadata.Name = $optionSetName
$optionSetMetadata.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($optionSetDisplayName, 1036)
$optionSetMetadata.IsGlobal = $true
$optionSetMetadata.OptionSetType = [Microsoft.Xrm.Sdk.Metadata.OptionSetType]::Picklist
$optionSetMetadata.Options.Add($firstOption)

$createRequest = New-Object Microsoft.Xrm.Sdk.Messages.CreateOptionSetRequest
$createRequest.OptionSet = $optionSetMetadata
$createRequest.SolutionUniqueName = $solutionUniqueName

try {
    $response = $conn.Execute($createRequest)
    Write-Host "OptionSet '$optionSetName' cree." -ForegroundColor Green
}
catch {
    Write-Host "Erreur creation : $_" -ForegroundColor Red
    Write-Host "Si l'OptionSet existe deja, supprime-le d'abord." -ForegroundColor Yellow
    exit
}

# ============================================================
# ETAPE 6 : AJOUT DES OPTIONS UNE PAR UNE
# ============================================================
Write-Host ""
Write-Host "ETAPE 6 : Ajout des departements..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

$total = $departements.Count
$current = 0

# On saute le premier (deja ajoute)
for ($i = 1; $i -lt $departements.Count; $i++) {
    $dept = $departements[$i]
    $current++
    
    $insertRequest = New-Object Microsoft.Xrm.Sdk.Messages.InsertOptionValueRequest
    $insertRequest.OptionSetName = $optionSetName
    $insertRequest.Value = $dept.Value
    $insertRequest.Label = New-Object Microsoft.Xrm.Sdk.Label($dept.Label, 1036)
    $insertRequest.SolutionUniqueName = $solutionUniqueName
    
    try {
        $conn.Execute($insertRequest) | Out-Null
        Write-Host "  [$current/$($total-1)] $($dept.Label)" -ForegroundColor Gray
    }
    catch {
        Write-Host "  Erreur sur $($dept.Label) : $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "$total departements ajoutes." -ForegroundColor Green

# ============================================================
# ETAPE 7 : PUBLICATION
# ============================================================
Write-Host ""
Write-Host "ETAPE 7 : Publication..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

try {
    Publish-CrmAllCustomization -conn $conn
    Write-Host ""
    Write-Host "Publie !" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Publication automatique echouee." -ForegroundColor Yellow
    Write-Host "Publie manuellement dans Power Apps : Solution > Publier toutes les personnalisations" -ForegroundColor Yellow
}
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  TERMINE" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "OptionSet '$optionSetName' disponible dans ta solution." -ForegroundColor White
Write-Host ""