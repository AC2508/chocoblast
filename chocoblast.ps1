Start-Sleep -Seconds 3
# Fonction pour vérifier si un processus est en cours d'exécution
function Get-ProcessByName {
    param (
        [string]$processName
    )
    return Get-Process | Where-Object { $_.Name -eq $processName } | Select-Object -First 1
}

# Vérifier si Outlook est déjà en cours d'exécution
$outlookRunning = Get-ProcessByName -processName "OUTLOOK"
$outlookWasRunning = $false

if ($null -ne $outlookRunning) {
    $outlookWasRunning = $true
}

try {
    # Créer une instance de l'application Outlook
    $outlook = New-Object -ComObject Outlook.Application

    # Récupérer le namespace MAPI
    $namespace = $outlook.GetNamespace("MAPI")

    # Initialiser la variable pour l'adresse e-mail nominative
    $nominativeAddress = $null

    # Parcourir tous les comptes pour trouver l'adresse nominative
    foreach ($account in $namespace.Accounts) {
        if ($account.SmtpAddress -match "aps-si.com" -and $account.SmtpAddress -ne "support@aps-si.com") {
            $nominativeAddress = $account.SmtpAddress
            break
        }
    }

    if ($nominativeAddress) {
        # Créer un nouvel e-mail
        $email = $outlook.CreateItem(0)

        # Définir les propriétés de l'e-mail
        $email.Subject = "Chocoblast"
        $email.Body = "La prochaine tournee de pains au chocolat est pour moi"
        
        # Ajouter plusieurs destinataires en les séparant par un point-virgule
        $email.To = "alexandre.corbineau@aps-si.com; dominique.neves@aps-si.com; sofiane.lachi@aps-si.com; jean-marie.corbineau@aps-si.com; matteo.auguet@aps-si.com; idriss.bopda@aps-si.com; yannis.filali@aps-si.com"
        
        # Définir l'adresse e-mail de l'expéditeur
        $email.SentOnBehalfOfName = $nominativeAddress
        
        # (Optionnel) Ajouter une pièce jointe
        # $email.Attachments.Add("chemin\vers\la\pièce\jointe.txt")
        
        # Envoyer l'e-mail
        $email.Send()
        
        Write-Host "Email envoye avec succès depuis $nominativeAddress"
    } else {
        Write-Host "Erreur : Aucune adresse nominative trouvee"
    }
} catch {
    Write-Host "Erreur lors de l'envoi de l'email : $_"
} finally {
    # Fermer Outlook si nous l'avons ouvert
    if (-not $outlookWasRunning -and $null -ne $outlook) {
        $outlook.Quit()
        Write-Host "Outlook fermé car il a été ouvert par le script."
    }
     
}

# Garder la fenêtre ouverte
Read-Host -Prompt "Appuyez sur Entree pour fermer cette fenetre"
