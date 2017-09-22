Function f_InfoPrompt($Message) {
    $value = ""
    [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
    $value = [Microsoft.VisualBasic.Interaction]::InputBox($Message, "Data Prompt")
    Return $value 
}

#######################

Function f_Output_Log ($string) {
	Write-Host ($string) -foregroundcolor Yellow
}
#######################
Function f_Get_Credential_Stored($user, $pwd) {
    $MyCredentials=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $user, ($pwd | ConvertTo-SecureString)
    Return $MyCredentials
}