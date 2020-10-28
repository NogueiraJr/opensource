#variaveis
$versao          = "v1.5"
$tabela          = "TABELA_TMP01"
$notificaLimite  = 5000

$newFilesPath    = "C:\newFile\"
$pathProcess     = "C:\process\"

$fileNameProcess = "Coletor"
$extensionFile   = ".xlsx"
$newFile         = [string]::Empty

$emailsDestino   = "seu@email.com"
$scriptName      = "coletor.$versao"

#functions
function logIni() {
    "" | Out-File ($pathProcess + "trace.log") -Append
    "Sistema Automatico" | Out-File ($pathProcess + "trace.log") -Append
    "PowerShell Windows" | Out-File ($pathProcess + "trace.log") -Append
    "Processo Coletor [$tabela] $versao" | Out-File ($pathProcess + "trace.log") -Append
    "" | Out-File ($pathProcess + "trace.log") -Append
}

function WriteHost([string]$ref, [string]$msg) {
    if ($ref -eq "1") { SendMail($msg) }
    if ($msg.Length -gt 0) {
        $agoraFormat = ((Get-Date).ToString('dd/MM/yyyy HH:mm:ss'))
        $log = ($agoraFormat + " " + $msg)
    }
    $log + " " | Out-File ($pathProcess + "trace.log") -Append
}

function SendMail($msgMail) {
    try {
        $proxy = New-WebServiceProxy -Uri http://server/oper/SendMail.asmx?WSDL
        $retorno = $proxy.Envia("seu@email.com", 
            $emailsDestino, 
            $scriptName, 
            $msgMail, "")
    } catch { 
        WriteHost "0" "Problemas enviando o e-mail notificador"
        WriteHost "0" "ERRO: $($_.Exception.ToString())"
    } finally { $proxy.dispose() }
}

function Main() {
    logIni
    WriteHost 0 "Iniciado"
    WriteHost 1 "Lendo Planilha [$newFile]"
    
    ls -Path $pathProcess -Filter *.dat | Remove-Item -Force

    $file = ($pathProcess + $fileNameProcess)
    
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $wb = $Excel.Workbooks.Open($file)
    #foreach ($ws in $wb.Worksheets)
    #{
        #$ws.SaveAs(("$file.dat"), 6)
    #}
    $wb.SaveAs(("$file.dat"), 6, $null, $null, $null, $true)
    
    foreach ($line in $wb.Worksheets.Item(1).Rows)
    {
        $line
    }    

    $wb.Close();
    $Excel.Quit()

    $qtdePlanilha = ((gc ($file + ".dat")).Count-1)
    WriteHost 0 "Formatando os dados"
    (gc ($file + ".dat"))              -replace(',', ';') > ($file + "_pontoVirgula.dat")
    WriteHost 0 "Aplicando 300 5"
    (gc ($file + "_pontoVirgula.dat")) -replace(';300$', ';5') > ($file + "300_5.dat")
    WriteHost 0 "Aplicando 70000 6"
    (gc ($file + "300_5.dat"))         -replace(';70000$', ';6') > ($file + "70000_6.dat")
    WriteHost 0 "Preparando arquivo final"
    (gc ($file + "70000_6.dat") | Select-Object -Skip 1) > ($file + "ToOracle.dat")

    WriteHost 0 "Conectando no Oracle"
    $Assembly = [System.Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient")

    if ( $Assembly ) { }
    else {
    	WriteHost 1 "System.Data.OracleClient nao carregado! Saindo..."
    	Exit 1
    }

    $OracleConnectionString = "Data Source=servidorAqui;User Id=usuarioAqui;Password=senhaAqui"

    $OracleConnection = New-Object System.Data.OracleClient.OracleConnection($OracleConnectionString);
    try {
        $OracleConnection.Open()
        }
    catch {
    	WriteHost 1 "Problemas conectando no Oracle! Saindo..."
        WriteHost 1 "Dump : $($_.Exception.ToString())"
        Exit 1
        }
    finally { }

    try {
        WriteHost 0 "Backup dos dados atuais do Coletor no Oracle"
        $OracleSQLQuery = "SEU SELECT NA TABELA AQUI" #(gc ($path + "script.bkp.sql"))
        $SelectCommand = New-Object System.Data.OracleClient.OracleCommand;
        $SelectCommand.Connection = $OracleConnection
        $SelectCommand.CommandText = $OracleSQLQuery
        $SelectCommand.CommandType = [System.Data.CommandType]::Text

        $SelectDataTable = New-Object System.Data.DataTable
        $SelectDataTable.Load($SelectCommand.ExecuteReader())
        
        $fileBkp = ($file + $agora + ".bkp")
        $SelectDataTable > $fileBkp
        WriteHost 0 "Backup realizado [$fileBkp]"
    }
    catch {
    	WriteHost 1 "[$OracleSQLQuery]"
    	WriteHost 1 "Erro: $($_.Exception.ToString())"
    	Exit 1
    }
    finally { $SelectCommand.Dispose() }

    try {
    	$HostUpdateCommand = New-Object System.Data.OracleClient.OracleCommand;
    	$HostUpdateCommand.Connection = $OracleConnection
    	$HostUpdateCommand.CommandType = [System.Data.CommandType]::Text
        
        WriteHost 0 "Apagando dados atuais do Coletor no Oracle"
    	$HostUpdateHostSQL = "TRUNCATE TABLE $tabela"
    	$HostUpdateCommand.CommandText = $HostUpdateHostSQL
    	$HostUpdateCommand.ExecuteNonQuery() | Out-Null

        WriteHost 0 "Inserindo novos dados do Coletor no Oracle"
        $contentFile = (gc ($file + "ToOracle.dat"))
        WriteHost 0 "Serao inseridos " + $contentFile.Count + " novos itens"
        $contagem = 0
        $notifica = $notificaLimite
        foreach ($lineRead in $contentFile) {
            $dados = $lineRead.Split(";")
        	$HostUpdateHostSQL = "INSERT INTO $tabela VALUES (" + $dados[0] + ", " + $dados[1] + ", " + $dados[2] + ", " + $dados[3] + ", " + $dados[4] + ")"
        	$HostUpdateCommand.CommandText = $HostUpdateHostSQL
        	$HostUpdateCommand.ExecuteNonQuery() | Out-Null
            if ($contagem -ge $notifica) {
                WriteHost 0 "Ate o momento inseridos " + $contagem
                $notifica += $notificaLimite
            }
            $contagem++
        }
        if ($qtdePlanilha -eq $contagem) {
            WriteHost 0 "Sucesso!"
            WriteHost 0 "Qtde Planilha = Qtde Oracle: $contagem itens"
        } else {
            WriteHost 1 ">> Algo saiu errado!"
            WriteHost 1 ">> A quantidade da planilha ($qtdePlanilha) eh diferente da quantidade inserida no Oracle ($contagem)"
        }
        }
    catch {
    	WriteHost 1 "[$HostUpdateHostSQL]"
    	WriteHost 1 "Erro: $($_.Exception.ToString())"
    	Exit 1
    }
    finally { $HostUpdateCommand.Dispose() }

    WriteHost 0 "Criando historico"
    WriteHost 1 "Finalizado"
    $agora = ((Get-Date).ToString('yyyyMMddHHmmss'))
    mkdir ($pathProcess + "$agora")      | out-null
    ls -Path $pathProcess -Filter *.log  | Move-Item -Destination ($pathProcess + "$agora")
    ls -Path $pathProcess -Filter *.dat  | Move-Item -Destination ($pathProcess + "$agora")
    ls -Path $pathProcess -Filter *.bkp  | Move-Item -Destination ($pathProcess + "$agora")
    ls -Path $pathProcess -Filter *.xlsx | Move-Item -Destination ($pathProcess + "$agora")
}

while (1 -eq 1) {
    $newFiles = gci -Path $newFilesPath -Filter *.xls* | Select-Object -ExpandProperty Name | sort -Property LastWriteTime
    if ($newFiles.count -gt 0) {
        if ($newFiles.count -eq 1) {
            $newFile = $newFiles
        } else {
            $newFile = $newFiles[0]
        }
        ls -Path $newFilesPath -Filter $newFile | Move-Item -Destination ($pathProcess + $fileNameProcess + $extensionFile) -Force
        #Start-Sleep -Seconds 10
        Main
    } else {Exit 1}
}
