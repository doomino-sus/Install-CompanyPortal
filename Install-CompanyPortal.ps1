# Skrypt instalacji Microsoft Company Portal z automatycznym logowaniem
# Wymaga uruchomienia jako Administrator

param(
    [Parameter(Mandatory=$true)]
    [string]$Username,
    
    [Parameter(Mandatory=$true)]
    [SecureString]$Password
)

# Funkcja do wyświetlania kolorowych komunikatów
function Write-ColoredOutput {
    param([string]$Message, [string]$Color = "Green")
    Write-Host $Message -ForegroundColor $Color
}

# Sprawdzenie uprawnień administratora
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-ColoredOutput "BŁĄD: Skrypt wymaga uprawnień administratora!" "Red"
    exit 1
}

Write-ColoredOutput "=== INSTALACJA MICROSOFT COMPANY PORTAL ===" "Cyan"
Write-ColoredOutput "Użytkownik: $Username" "Yellow"

# Instalacja Company Portal
Write-ColoredOutput "1. Instalowanie Microsoft Company Portal..." "Green"

try {
    # Próba instalacji przez winget
    $wingetCheck = Get-Command winget -ErrorAction SilentlyContinue
    if ($wingetCheck) {
        Write-ColoredOutput "Używam winget do instalacji..." "Yellow"
        winget install --id=9WZDNCRFJ3PZ --source=msstore --accept-package-agreements --accept-source-agreements
    } else {
        # Alternatywnie przez PowerShell i Microsoft Store
        Write-ColoredOutput "Instaluję przez Microsoft Store..." "Yellow"
        Start-Process "ms-windows-store://pdp/?productid=9WZDNCRFJ3PZ"
        Write-ColoredOutput "Otwieram Microsoft Store - proszę dokończyć instalację ręcznie" "Yellow"
        Read-Host "Naciśnij Enter po zakończeniu instalacji"
    }
} catch {
    Write-ColoredOutput "Błąd podczas instalacji: $($_.Exception.Message)" "Red"
    exit 1
}

# Oczekiwanie na instalację
Write-ColoredOutput "2. Oczekiwanie na zakończenie instalacji..." "Green"
Start-Sleep -Seconds 10

# Sprawdzenie czy aplikacja jest zainstalowana
$companyPortalPath = Get-AppxPackage -Name "*CompanyPortal*" | Select-Object -First 1
if (-not $companyPortalPath) {
    Write-ColoredOutput "BŁĄD: Company Portal nie został znaleziony!" "Red"
    exit 1
}

Write-ColoredOutput "3. Company Portal został pomyślnie zainstalowany!" "Green"

# Uruchomienie Company Portal
Write-ColoredOutput "4. Uruchamianie Company Portal..." "Green"
Start-Process "ms-windows-store://launch/?PfamID=Microsoft.CompanyPortal_8wekyb3d8bbwe"

# Oczekiwanie na uruchomienie aplikacji
Start-Sleep -Seconds 5

# Skrypt automatyzacji logowania (używa UI Automation)
Write-ColoredOutput "5. Rozpoczynam automatyczne logowanie..." "Green"

# Konwersja SecureString na plain text (tylko dla procesu logowania)
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
$PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# Automatyzacja UI - wymaga modułu UIAutomation lub alternatywne rozwiązanie
try {
    Add-Type -AssemblyName System.Windows.Forms
    
    # Oczekiwanie na okno logowania (może być różne w zależności od wersji)
    Write-ColoredOutput "Szukam okna logowania..." "Yellow"
    Start-Sleep -Seconds 3
    
    # Wysyłanie danych logowania
    [System.Windows.Forms.SendKeys]::SendWait($Username)
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait($PlainPassword)
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    
    Write-ColoredOutput "Dane logowania zostały wprowadzone automatycznie!" "Green"
    
} catch {
    Write-ColoredOutput "Automatyczne logowanie nie powiodło się. Proszę zalogować się ręcznie." "Yellow"
    Write-ColoredOutput "Użyj danych: $Username" "Yellow"
}

# Oczyszczenie hasła z pamięci
$PlainPassword = $null
[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

Write-ColoredOutput "=== INSTALACJA ZAKOŃCZONA ===" "Cyan"
Write-ColoredOutput "Company Portal jest uruchomiony. Sprawdź czy logowanie przebiegło pomyślnie." "Green"

# Opcjonalne: Sprawdzenie stanu rejestracji w Intune
Write-ColoredOutput "6. Sprawdzanie stanu rejestracji..." "Green"
try {
    $enrollmentStatus = Get-WmiObject -Namespace "root\cimv2\mdm\dmmap" -Class "MDM_DevDetail_Ext01" -ErrorAction SilentlyContinue
    if ($enrollmentStatus) {
        Write-ColoredOutput "Urządzenie jest zarejestrowane w Intune!" "Green"
    } else {
        Write-ColoredOutput "Sprawdź status rejestracji w aplikacji Company Portal" "Yellow"
    }
} catch {
    Write-ColoredOutput "Nie można sprawdzić stanu rejestracji automatycznie" "Yellow"
}
