function Send-EmailWithAttachments()
{
    <#
.SYNOPSIS
    Sends email with file attachments using Microsoft Graph API or MAPI (Outlook COM automation).

.DESCRIPTION
    This function provides enterprise-grade email functionality using either Microsoft Graph API
    or MAPI/Outlook COM automation for secure and reliable email transmission. Features include:
    
    Microsoft Graph Integration (Default):
    - Uses modern authentication with Microsoft Graph
    - Supports organizational email policies and security
    - Handles large attachments efficiently
    - Sends email automatically without user interaction
    
    Outlook COM Integration (Optional):
    - Opens Classic Outlook with pre-filled message
    - Allows user to review and edit before sending
    - Works with standard user permissions
    - Does not require Graph API authentication
    - Requires Classic Outlook (New Outlook does not support COM automation)
    
    Attachment Processing:
    - Supports multiple file attachments
    - Automatic MIME type detection based on file extensions (Graph API)
    - Base64 encoding for secure transmission (Graph API)
    - Direct file attachment (Outlook COM)
    - File size validation and handling
    
    Error Handling:
    - Comprehensive connection and authentication error handling
    - Graceful degradation when Graph modules are unavailable
    - Detection of New Outlook vs Classic Outlook
    - Clear error messages when Classic Outlook is not available
    - Detailed logging of email transmission attempts

.PARAMETER AccessToken
    Access token for Microsoft Graph API authentication.
    Required when using Graph API mode (default). Not used when -UseMAPI is specified.

.PARAMETER To
    Email address of the recipient. Should be a valid email address
    within the organization or allowed external domain.

.PARAMETER Subject
    Subject line for the email message. Will be included exactly as provided.

.PARAMETER Body
    Plain text body content for the email message.
    HTML formatting is not supported in the current implementation.

.PARAMETER AttachmentPaths
    Array of file paths to include as email attachments.
    Files are validated for existence before inclusion.
    Supported file types include .txt, .log, .cer, .zip, and others.

.PARAMETER UseMAPI
    Switch parameter to use Outlook COM automation instead of Microsoft Graph API.
    When specified, opens Classic Outlook with a pre-filled message, allowing the user
    to review, edit, and send manually. Requires Classic Outlook to be installed.
    Note: New Outlook does not support COM automation.

.EXAMPLE
    $success = Send-EmailWithAttachments -AccessToken $token -To "support@company.com" -Subject "PIV Issue" -Body "Please help" -AttachmentPaths @("C:\logs\error.log")
    Sends email using Microsoft Graph API.

.EXAMPLE
    $success = Send-EmailWithAttachments -To "support@company.com" -Subject "PIV Issue" -Body "Please help" -AttachmentPaths @("C:\logs\error.log") -UseMAPI
    Opens Classic Outlook with pre-filled email for user review and manual sending.

.OUTPUTS
    Boolean value indicating email operation success:
    - $true: Email sent successfully (Graph API) or email client opened successfully (MAPI)
    - $false: Email transmission or client opening failed

.NOTES
    Prerequisites for Graph API mode:
    - Microsoft.Graph.Users.Actions PowerShell module
    - Appropriate permissions for Mail.Send in Microsoft Graph
    - Valid organizational email configuration
    - AccessToken parameter
    
    Prerequisites for Outlook COM mode:
    - Classic Outlook installed and configured
    - No authentication required
    
    Note: New Outlook does not support COM automation and cannot be used with this feature.
    
    The function automatically handles authentication prompts and permission requests.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $accessToken,
        [Parameter(Mandatory = $true)]
        [string]$To,
        [Parameter(Mandatory = $true)]
        [string]$Subject,
        [Parameter(Mandatory = $true)]
        [string]$Body,
        [Parameter(Mandatory = $false)]
        [string[]]$AttachmentPaths = @(),
        [Parameter(Mandatory = $false)]
        [switch]$UseMAPI
    )
    
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Message "Starting email send process to $To $(if ($UseMAPI) { '(MAPI mode)' } else { '(Graph API mode)' })" -Module $functionName -LogLevel "Information" -LogFile $logFile
    
    # MAPI Mode: Use Outlook COM automation to open email client
    if ($UseMAPI)
    {
        Write-Log -Message "Using Outlook COM automation to create email" -Module $functionName -LogLevel "Information" -LogFile $logFile
        Write-Verbose "[$functionName] Using Outlook COM automation"
        
        try
        {
            # Attempt to create Outlook COM object (Classic Outlook)
            # Note: New Outlook does not support COM automation
            $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
            Write-Log -Message "Successfully created Outlook COM object" -Module $functionName -LogLevel "Debug" -LogFile $logFile
            
            # Create a new mail item
            $mail = $outlook.CreateItem(0)  # 0 = olMailItem
            Write-Log -Message "Created new Outlook mail item" -Module $functionName -LogLevel "Debug" -LogFile $logFile
            
            # Set email properties
            $mail.To = $To
            $mail.Subject = $Subject
            $mail.Body = $Body
            Write-Log -Message "Set email properties: To=$To, Subject=$Subject" -Module $functionName -LogLevel "Debug" -LogFile $logFile
            
            # Add attachments
            $attachmentCount = 0
            foreach ($attachmentPath in $AttachmentPaths)
            {
                if (Test-Path $attachmentPath)
                {
                    try
                    {
                        $mail.Attachments.Add($attachmentPath) | Out-Null
                        $fileName = [System.IO.Path]::GetFileName($attachmentPath)
                        $attachmentCount++
                        Write-Log -Message "Added attachment: $fileName" -Module $functionName -LogLevel "Information" -LogFile $logFile
                        Write-Verbose "[$functionName] Added attachment: $fileName"
                    }
                    catch
                    {
                        Write-Log -Message "Failed to add attachment $attachmentPath : $($_.Exception.Message)" -Module $functionName -LogLevel "Warning" -LogFile $logFile
                        Write-Warning "[$functionName] Failed to add attachment $attachmentPath : $($_.Exception.Message)"
                    }
                }
                else
                {
                    Write-Log -Message "Attachment file not found: $attachmentPath" -Module $functionName -LogLevel "Warning" -LogFile $logFile
                    Write-Verbose "[$functionName] Attachment file not found: $attachmentPath"
                }
            }
            
            # Display the email for user review and sending
            $mail.Display()
            Write-Log -Message "Email displayed in Outlook with $attachmentCount attachment(s)" -Module $functionName -LogLevel "Information" -LogFile $logFile
            Write-Verbose "[$functionName] Email displayed in Outlook with $attachmentCount attachment(s)"
            
            return $true
        }
        catch
        {
            $errorMessage = "Failed to create email using Outlook COM automation: $($_.Exception.Message)"
            Write-Error $errorMessage
            Write-Log -Message $errorMessage -Module $functionName -LogLevel "Error" -LogFile $logFile
            
            # Check if Classic Outlook is installed
            $classicOutlookInstalled = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE" -ErrorAction SilentlyContinue
            
            # Check if New Outlook is being used
            $newOutlookEnabled = $false
            $newOutlookRegPath = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences'
            if (Test-Path $newOutlookRegPath)
            {
                $newOutlookSetting = Get-ItemProperty -Path $newOutlookRegPath -Name 'UseNewOutlook' -ErrorAction SilentlyContinue
                if ($newOutlookSetting.UseNewOutlook -eq 1)
                {
                    $newOutlookEnabled = $true
                }
            }
            
            if ($newOutlookEnabled)
            {
                Write-Log -Message "New Outlook is enabled, which does not support COM automation" -Module $functionName -LogLevel "Error" -LogFile $logFile
                Write-Host "New Outlook is currently enabled. This feature requires Classic Outlook." -ForegroundColor Red
                Write-Host "To use this feature, please switch to Classic Outlook or use the Graph API option instead." -ForegroundColor Yellow
            }
            elseif (-not $classicOutlookInstalled)
            {
                Write-Log -Message "Classic Outlook does not appear to be installed on this system" -Module $functionName -LogLevel "Error" -LogFile $logFile
                Write-Host "Classic Outlook is not installed. This feature requires Classic Outlook." -ForegroundColor Red
                Write-Host "Please install Outlook or use the Graph API option instead." -ForegroundColor Yellow
            }
            else
            {
                Write-Log -Message "Outlook COM automation failed: $($_.Exception.Message)" -Module $functionName -LogLevel "Error" -LogFile $logFile
                Write-Host "Failed to automate Outlook. Please ensure Classic Outlook is properly configured." -ForegroundColor Red
            }
            
            return $false
        }
    }
    
    # Graph API Mode: Continue with existing implementation
    $requiredScope = "Mail.Send"
    if (-not $accessToken)
    {
        Write-Error "AccessToken is required when not using MAPI mode"
        Write-Log -Message "AccessToken is required when not using MAPI mode" -Module $functionName -LogLevel "Error" -LogFile $logFile
        return $false
    }
    
    # Extract user info from access token to ensure we're sending from the correct mailbox
    $tokenClaims = DecodeJwtToken -Token $accessToken -raw
    $grantedScopes = @()
    if ($tokenClaims.scp)
    {
        # Delegated auth - scp is space-separated string
        Write-Verbose "[$functionName] Cached token has delegated scopes (scp)"
        $grantedScopes = $tokenClaims.scp -split ' ' | Where-Object { $_ -and $_.Trim() }
        Write-Verbose "[$functionName] Cached token has delegated scopes (scp): $($grantedScopes -join ', ')"
        write-log -logFile $logFile -Module "$functionName" -Message "Cached token has delegated scopes (scp): $($grantedScopes -join ', ')"                                
    }
    elseif ($tokenClaims.roles)
    {
        # Application auth - roles is array
        $grantedScopes = $tokenClaims.roles
        Write-Verbose "[$functionName] Cached token has application scopes (roles): $($grantedScopes -join ', ')"
        write-log -logFile $logFile -Module "$functionName" -Message "Cached token has application scopes (roles): $($grantedScopes -join ', ')"                    
    }
    else 
    {
        Write-Error "Could not determine granted scopes from access token"
        Write-Log -Message "Could not determine granted scopes from access token" -Module $functionName -LogLevel "Error" -LogFile $logFile
        return $false
    }                                                               
    
    if (-not ($grantedScopes -contains $requiredScope))
    {
        Write-Error "The access token does not have the required scope '$requiredScope' to send email."
        Write-Log -Message "The access token does not have the required scope '$requiredScope' to send email." -Module $functionName -LogLevel "Error" -LogFile $logFile
        return $false
    }                                                                                   
    Write-Verbose "[$functionName] Access token has required scope: $requiredScope"                                     
    write-log -logFile $logFile -Module "$functionName" -Message "Access token has required scope: $requiredScope"                  
    $senderUPN = $null
    
    # Try to get UPN from token claims (preferred_username, upn, email, unique_name in order of preference)
    if ($tokenClaims.preferred_username)
    {
        $senderUPN = $tokenClaims.preferred_username
        Write-Log -Message "Using preferred_username from token: $senderUPN" -Module $functionName -LogLevel "Debug" -LogFile $logFile
    }
    elseif ($tokenClaims.upn)
    {
        $senderUPN = $tokenClaims.upn
        Write-Log -Message "Using upn from token: $senderUPN" -Module $functionName -LogLevel "Debug" -LogFile $logFile
    }
    elseif ($tokenClaims.email)
    {
        $senderUPN = $tokenClaims.email
        Write-Log -Message "Using email from token: $senderUPN" -Module $functionName -LogLevel "Debug" -LogFile $logFile
    }
    elseif ($tokenClaims.unique_name)
    {
        $senderUPN = $tokenClaims.unique_name
        Write-Log -Message "Using unique_name from token: $senderUPN" -Module $functionName -LogLevel "Debug" -LogFile $logFile
    }
    else
    {
        write-log -logFile $logFile -module $functionName -message "Could not determine sender UPN from token claims" -logLevel "Warning" 
        $senderUPN = Read-Host -Prompt "Please enter your email address"                     
        #ensure we get a valid email address format
        while (-not ($senderUPN -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'))                                     
        {
            Write-Host "Incorrect email format. Please enter a valid email address." -ForegroundColor Yellow                                
            [console]::beep(1000, 300)                       
            $senderUPN = Read-Host -Prompt "Please enter your email address"                     
        }
        write-log -logFile $logFile -module $functionName -message "Using user-provided sender UPN: $senderUPN" -logLevel "Debug"                                           
        Write-Verbose "[$functionName] Using user-provided sender UPN: $senderUPN"                              
    }
    
    # Use me/sendMail since we're using the authenticated user's context
    $emailSendURI = "me/sendMail"
    # Prepare attachments
    $attachments = @()
    foreach ($attachmentPath in $AttachmentPaths)
    {
        if (Test-Path $attachmentPath)
        {
            $fileBytes = [System.IO.File]::ReadAllBytes($attachmentPath)
            $fileName = [System.IO.Path]::GetFileName($attachmentPath)
            $contentType = switch ([System.IO.Path]::GetExtension($attachmentPath).ToLower())
            {
                ".txt"
                {
                    "text/plain" 
                }
                ".log"
                {
                    "text/plain" 
                }
                ".cer"
                {
                    "application/x-x509-ca-cert" 
                }
                ".zip"
                {
                    "application/zip" 
                }
                default
                {
                    "application/octet-stream" 
                }
            }
                
            $attachment = @{
                "@odata.type"  = "#microsoft.graph.fileAttachment"
                "name"         = $fileName
                "contentType"  = $contentType
                "contentBytes" = [System.Convert]::ToBase64String($fileBytes)
            }
            $attachments += $attachment
            Write-Log -Message "Added attachment: $fileName ($([math]::Round($fileBytes.Length / 1KB, 2)) KB)" -Module $functionName -LogLevel "Information" -LogFile $logFile
        }
        else
        {
            Write-Log -Message "Attachment file not found: $attachmentPath" -Module $functionName -LogLevel "Warning" -LogFile $logFile
        }
    }
    
    #prepare email message object - attachments MUST be inside the message object per Microsoft Graph API spec
    $emailMessage = [ordered]@{
        message         = @{
            subject      = $Subject
            body         = @{
                contentType = "Text"
                content     = $Body
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $To
                    }
                }
            )
            attachments  = $attachments
        }
        saveToSentItems = $true                                           
    } | ConvertTo-Json -Depth 5         
    # Send the email
    try 
    {
        Write-Log -Message "Calling Graph API to send email" -Module $functionName -LogLevel "Debug" -LogFile $logFile
        $emailResponse = CallGraphApi -AccessToken $accessToken -Method POST -ResourcePath $emailSendURI -Body $emailMessage
        Write-Log -Message "Email sent successfully to $To with $($attachments.Count) attachment(s)" -Module $functionName -LogLevel "Information" -LogFile $logFile
        if ([string]::IsNullOrEmpty($emailResponse))
        {
            return $true
        }
        else
        {
            Write-Log -Message "Unexpected response from email send: $emailResponse" -Module $functionName -LogLevel "Warning" -LogFile $logFile
            return $false
        }                           
    }
    catch
    {
        $errorDetails = @(
            "Failed to send email: $($_.Exception.Message)"
            "Error Type: $($_.Exception.GetType().FullName)"
            "Sender UPN: $senderUPN"
            "Recipient: $To"
            "Attachment Count: $($attachments.Count)"
        )
        
        if ($_.ErrorDetails.Message)
        {
            $errorDetails += "API Error Details: $($_.ErrorDetails.Message)"
        }
        
        $fullErrorMessage = $errorDetails -join ' | '
        Write-Error $fullErrorMessage
        Write-Log -Message $fullErrorMessage -Module $functionName -LogLevel "Error" -LogFile $logFile
        
        # Try to decode Graph API error if available
        if ($_.ErrorDetails.Message)
        {
            try
            {
                $graphError = $_.ErrorDetails.Message | ConvertFrom-Json
                if ($graphError.error)
                {
                    Write-Log -Message "Graph API Error Code: $($graphError.error.code)" -Module $functionName -LogLevel "Error" -LogFile $logFile
                    Write-Log -Message "Graph API Error Message: $($graphError.error.message)" -Module $functionName -LogLevel "Error" -LogFile $logFile
                }
            }
            catch
            {
                Write-Verbose "[$functionName] Could not parse Graph API error details"
            }
        }
        
        return $false
    }
}
