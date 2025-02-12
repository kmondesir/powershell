 <#    
        .SYNOPSIS
            Creates an listener for http requests. 
        .DESCRIPTION
            Creates a listener for http requests to test applications 
        .PARAMETER Port
            Optional. Port to connect on.
        .EXAMPLE
            .\Start-APIServer.ps1
    #>
    [cmdletbinding()]
    param 
    (
        [Parameter(HelpMessage = 'Port number to listen on', Mandatory = $false, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]  
        [int]$port = 5000
    )
    Write-Host "Starting API service..."

    try {
        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add("http://+:$port/")
        $listener.Start()
        while ($listener.IsListening) {
            $message = ""
            $context = $listener.GetContext()
            $request = $context.Request
            $response = $context.Response
            $path = $request.Url.AbsolutePath

            Write-Host "Request received: $path"

            # Set CORS Headers
            $response.Headers.Add("Access-Control-Allow-Origin", "*")
            $response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            $response.Headers.Add("Access-Control-Allow-Headers", "Content-Type")
            $response.StatusCode = 200

            if ($request.HttpMethod -eq "OPTIONS") {
                $response.Close()
                continue
            }

            switch ($path) {
                "/api/health" {
                    $message = "API is running!"
                }
                default {
                    $response.StatusCode = 404
                    $message = "Not Found!"
                }
            }

            $responseBytes = [System.Text.Encoding]::UTF8.GetBytes($message)
            $response.OutputStream.Write($responseBytes, 0, $responseBytes.Length)
            $response.OutputStream.Close()
        }
    }
    catch {
        $message = $PSItem.Exception.Message
        Write-Host $message
    }
    finally {
        $listener.Stop()
        Write-Host "API service stopped"
    }
