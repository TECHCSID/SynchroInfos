# Recuperation du dossier CSID Update
$csidUpdatePath = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\CSiD\CSiDUpdate | Select-Object -ExpandProperty InstallLocation;
$paramguFile = Join-Path -Path $csidUpdatePath -ChildPath "paramgu.ini";
$SortieConfig = "";
# Récupération de l'heure d'installation
if (Test-Path $paramguFile -PathType leaf) {
    # Récupération de la ligne de configuration de la config ([ETUDE])
    $UpdateLine = Select-String -Path $paramguFile -Pattern "\[ETUDE\]";

    # Si on a trouvé un groupe [ETUDE]
    if ($null -ne $UpdateLine) {
        # Quatrième ligne à récupérer après [ETUDE] -- Application
        $ApplicationEtude = Get-Content $paramguFile | Select -Index ($UpdateLine.LineNumber + 2);
        
        if ($ApplicationEtude.StartsWith("Application") -and $ApplicationEtude.Contains('SynchroExch')) {
            $SortieConfig += "GU telecharge la synchro";
        } else {
            $SortieConfig += "GU ne telecharge pas la synchro";
        }
    }
    else {
        $SortieConfig += "Prolème de configuration GUpdate";
    }
}


# Recuperation du dossier de la synchro
$synchroPath = $null;
$synchroService = Get-WmiObject win32_service | ?{$_.Name -like 'Synchronisation iNot Exchange'};
if($null -ne $synchroService)
{
    $synchroPath = $synchroService.PathName.Trim().Trim('"');
    $SortieConfig += " / Synchro Version : ";
    $SortieConfig += [System.Diagnostics.FileVersionInfo]::GetVersionInfo($synchroPath).FileVersion;

    $SortieConfig += " / Etat du service : " + $synchroService.State;

    $synchroDirPath = [System.IO.Path]::GetDirectoryName($synchroPath);
    $ConfigSynchroPath = Join-Path -Path $synchroDirPath -ChildPath "Config"; 
    $ConfigFile = Join-Path -Path $ConfigSynchroPath -ChildPath "Config.xml"; 

    if(Test-Path -Path $ConfigFile)
    {
        [XML] $configXml = Get-Content -Path $ConfigFile -ErrorAction 'Stop';

        $ConfigOffice365 = $configXml.SelectSingleNode("//Office365");
        $ConfigNotamail = $configXml.SelectSingleNode("//Notamail");

        $SortieConfig += " / Configuration de la Synchro : ";

        if($ConfigOffice365.InnerText -eq $true)
        {
            $SortieConfig += "Office 365 Cloud";
        } elseif ($ConfigNotamail.InnerText -eq $true)
        {
            $SortieConfig += "Notamail";
        } else {
            $SortieConfig += "Exchange Local";
        }
    
        $configUserMapping = $configXml.SelectNodes("//UserMapping");

        if($configUserMapping.Count -ne 0)
        {
            if($configUserMapping.Count -eq 1) {
                $SortieConfig += " / 1 utilisateur trouve";
            } else {
                $SortieConfig += " / " + $configUserMapping.Count + " utilisateurs trouves";
            }
            $i = 1;
            foreach ($configUser in $configUserMapping)
            {
                $SortieConfig += " / Utilisateur $i : " + $configUser.PLoginName + " - Email : " + $configUser.XLogin;
                $i++;
            }
        } else {
            $SortieConfig += " / Pas d'utilisateurs"
        }
    }
} else {
    $SortieConfig = "Pas de Synchronisation Exchange";
}

Write-Host $SortieConfig;