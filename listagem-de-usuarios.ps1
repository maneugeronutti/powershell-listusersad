Write-Host "        /         \                ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
Write-Host "       (_\_@@@@@_/_)               ~~~~~~~~~~~~~~~~~  MATHEUZIN - v2.0 ~~~~~~~~~~~~~~~~~"
Write-Host "           (o o)                   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
Write-Host "            \ /----------\\        Script para consultas e exportações de usuários"
Write-Host "             O||        ||~        que NUNCA se logaram no Active Directory!"
Write-Host "              ||------/@||         ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
Write-Host "              ||        ||"     

Write-Host ""
Write-Host "O script irá ignorar usuários que iniciam com o nome _MODELO e que pertencem a OU - Metabase"

$loop = $true

while ($loop -eq $true){
    Write-Host ""
    Write-Host "1. Listar/Exportar usuários que NUNCA logaram (TODOS)"
    Write-Host "2. Listar/Exportar usuários que NUNCA logaram (POR FILTRO DE ANO/MÊS)"
    Write-Host ""
    Write-Host 'q. Sair'
    Write-Host '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
    $optionSelected = Read-Host -Prompt "Selecione a opção desejada (ou 'q' para sair)"
    
    $loopYear = $true

    switch ($optionSelected) {
        '1' {

        $isExport = Read-host -Prompt "Deseja exportar esta consulta? (S/N)"
        Write-Host ""

        if ($isExport -eq 'S'){
            $fileExport = Read-host -Prompt "Informe o local e nome do arquivo (Default: C:\Temp\listUsers.csv)"

            if ($fileExport -eq '') { $fileExport = "C:\Temp\listUsers.csv"}

            try {
                Get-ADUser -Filter * -Properties name,sAMAccountName,whenCreated,lastLogonDate,Office,Description | Where-Object {$_.lastLogonDate -eq $null -and $_.Name -notlike "_Modelo*" -and $_.DistinguishedName -notlike "*OU=Metabase,DC=induscar,DC=lan,DC=br"} | Select Name,sAMAccountName,whenCreated,Office,Description | Export-Csv $fileExport -NoTypeInformation
                Write-Host "";Write-Host ""
                Write-Host "Processo finalizado!"

                $returnMenu = Read-Host -Prompt "Desejar retornar ao menu principal? (S/N)"
                if ($returnMenu -eq 'S') { $loop = $true } elseif ($returnMenu -eq 'N') { Write-Host "Tchauzinho!";$loop = $false; exit 0 } else { Write-Host "Opção inválida!"; exit 0 }

            } catch {
                Write-Host ""; Write-Host ""
                Write-Host "ERRO! Problemas na execução do comando."
                $loop = $false
            }

        }elseif ($isExport -eq 'N'){
            try {
                Get-ADUser -Filter * -Properties name,sAMAccountName,whenCreated,lastLogonDate,Office,Description | Where-Object {$_.lastLogonDate -eq $null -and $_.Name -notlike "_Modelo*" -and $_.DistinguishedName -notlike "*OU=Metabase,DC=induscar,DC=lan,DC=br"} | Select Name,sAMAccountName,whenCreated,Office,Description

                Write-Host ""; Write-Host ""
                Write-Host "Processo finalizado!"

                $returnMenu = Read-Host -Prompt "Desejar retornar ao menu principal? (S/N)"
                if ($returnMenu -eq 'S') { $loop = $true } elseif ($returnMenu -eq 'N') { Write-Host "Tchauzinho!";$loop = $false; exit 0 } else { Write-Host "Opção inválida!"; exit 0 }
                

            } catch {
                Write-Host ""; Write-Host ""
                Write-Host "ERRO! Problemas na execução do comando."
                $loop = $false
            }
            
        }

        }
        '2' {
            while ($loopYear -eq $true) {
                $year = Read-host -Prompt "Digite o ano de corte"
                do {
                    $month = Read-Host "Digite o mês de corte"
                    $month = [int]$month
                    if (-not ($month -as [int])) {
                        Write-Host ""
                        Write-Host "Por favor, digite um valor numérico."
                    } elseif ($month -lt 1 -or $month -gt 12) {
                        Write-Host ""
                        Write-Host "Por favor, digite um valor entre 1 e 12."
                    } else {
                        $validInput = $true
                    }
                } while (-not $validInput)
                
                if ($month -ge 1 -and $month -le 12) {
                    
                    try {
                        Write-Host ""
                        $isExport = Read-host -Prompt "Deseja exportar esta consulta? (S/N)"
                        Write-Host ""

                        if ($isExport -eq 'S'){
                            $fileExport = Read-host -Prompt "Informe o local e nome do arquivo (Default: C:\Temp\listUsersByFilter.csv)"

                            if ($fileExport -eq '') { $fileExport = "C:\Temp\listUsersByFilter.csv"}
                            
                            Get-ADUser -Filter * -Properties name,sAMAccountName,whenCreated,lastLogonDate,Office,Description | Where-Object {$_.lastLogonDate -eq $null -and $_.whenCreated -lt "$year-$month-01T00:00:00.000Z" -and $_.Name -notlike "_Modelo*" -and $_.DistinguishedName -notlike "*OU=Metabase,DC=induscar,DC=lan,DC=br"} | Select Name,sAMAccountName,whenCreated,Office,Description | Export-Csv $fileExport -NoTypeInformation
                    
                            Write-Host ""; Write-Host ""
                            Write-Host "Processo finalizado!"
                            $returnMenu = Read-Host -Prompt "Desejar retornar ao menu principal? (S/N)"
                            if ($returnMenu -eq 'S') { $loop = $true;$loopYear = $false } elseif ($returnMenu -eq 'N') { Write-Host "Tchauzinho!";$loop = $false;$loopYear = $false; exit 0 } else { Write-Host "Opção inválida!"; exit 0 }
                        
                        } elseif ($isExport -eq 'N') {
                            try {
                                Get-ADUser -Filter * -Properties name,sAMAccountName,whenCreated,lastLogonDate,Office,Description | Where-Object {$_.lastLogonDate -eq $null -and $_.whenCreated -lt "$year-$month-01T00:00:00.000Z" -and $_.Name -notlike "_Modelo*" -and $_.DistinguishedName -notlike "*OU=Metabase,DC=induscar,DC=lan,DC=br"} | Select Name,sAMAccountName,whenCreated,Office,Description

                                Write-Host ""; Write-Host ""
                                Write-Host "Processo finalizado!"
                                $returnMenu = Read-Host -Prompt "Desejar retornar ao menu principal? (S/N)"
                                if ($returnMenu -eq 'S') { $loop = $true;$loopYear = $false } elseif ($returnMenu -eq 'N') { Write-Host "Tchauzinho!";$loop = $false;$loopYear = $false; exit 0 } else { Write-Host "Opção inválida!"; exit 0 }

                            } catch {
                            write-host $_
                                Write-Host ""; Write-Host ""
                                Write-Host "ERRO! Problemas na execução do comando."
                                $loop = $false
                            }
                            
                        }
                     
                    } catch {
                        Write-Host ""; Write-Host ""
                        Write-Host "ERRO! Problemas na execução do comando."
                        $loop = $false
                    }
                } else {
                    Write-Host ""; Write-Host ""
                    Write-Host "Valor de MÊS incorreto!"
                }
            }
            
        }
        'q' {
            Write-Host "Tchauzinho!"
            exit 0
        }
        default {
            Write-Host ''
            Write-Host "Opção '$optionSelected' inválida, tente novamente!"
        }
    }
}
