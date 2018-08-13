param (
    [Parameter(Mandatory=$true)][string]$source = '',
	[Parameter(Mandatory=$true)][string]$destination = '',
	[Parameter(Mandatory=$true)][string]$config_file = '',
	[Parameter(Mandatory=$true)][string]$service_name = '',
	[Parameter(Mandatory=$true)][string]$iis_site_name = '',
	[Parameter(Mandatory=$true)][string]$process_user = ''
)

trap #capture error and make sure it fails Bamboo
{
	"Error: $_" #output the error message
	exit 1 #exit with error code 1
}

#create folder if doesn't exist
if (-Not ("$destination\web-api"))
{
     md -path "$destination\web-api"
}

if (-Not ("$destination\agent-service"))
{
     md -path "$destination\agent-service"
}

#############################################################################################
# Deploy Agent																				#
#############################################################################################
$service = Get-Service -Name $service_name #retrieve the service


if($service.Status -eq "Running")
{
	Write-Output "Stopping Service: $service_name"
	$current_status = $service.Status #assign current status
	
	$process_id = 0 #initiate variable
	
	#get processes with the same .exe name
	$processes = Get-WmiObject -Class Win32_Process -Filter 'name like "Pasco.Middleware.Agent.exe"'
	
	#identify the process
	foreach($process in $processes)
	{
		$user = $process.getowner().user
		if($user -eq $process_user)
		{
			$process_id = $process.ProcessId #assign the proper id
			Write-Output "Process: Pasco.Middleware.Agent.exe, username: $process_user, id=$process_id."
		}
	}
	
	$temp = $service | Stop-Service #attempt to stop the service
	
	$attempt = 10
	do
	{
		$attempt--
		$service = Get-Service -Name $service_name
		$current_status = $service.Status
		Write-Output "Service Status: $current_status"
		
		$process = get-wmiobject -class win32_process -filter "ProcessId=$process_id"

		if($process -ne $Null)
		{
			Write-Output "Process is still running, wait for 10 seconds"
			sleep -Seconds 10 #wait for 10 seconds and check the status again
		}
		
	} until (($process -eq $Null) -or ($attempt -eq 0))
	
	if($process -ne $Null)
	{
		Write-Output "Process id=$process_id has not stopped after 5 minutes. Killing now"
		Stop-Process -id $process_id -Force
	}
	else
	{
		Write-Output "Process stopped successfully. Start copying files..."
	}
}
else
{
	Write-Output "Service is not running. Start copying files..."
}

sleep -Seconds 5 #in case any file is still locked

copy-item -path "$source\*" -destination "$destination\agent-service" -recurse -force
write-output "All files have been copied"

$agent_config_file = Get-Item -Path "$destination\agent-service\$config_file"
copy-item $agent_config_file -destination "$destination\agent-service\Pasco.Middleware.Agent.exe.config" -force
write-output "Copied $agent_config_file as Pasco.Middleware.Agent.exe.config"


#$config_files = Get-ChildItem "$destination\agent-service\*.config"
#$match_rule = '(?<=Server=MTWMID01;Database=)([^"]*)(?=;Trusted_Connection=True;)' #look up the database name
#foreach($this_config_file in $config_files)
#{
#    (Get-Content $this_config_file) -replace $match_rule, $database_name | Set-Content $this_config_file #replace database name to be what's specified in the parameter#
#	Write-Output "$config_file database name set"
#}

Write-Output "$service_name service is restarting..."
$temp = $service | Start-Service
$restarted_service_status = $service.Status


$check_service_attempt = 30
do
{
	$check_service_attempt--
	$service = Get-Service -Name $service_name
	$restarted_service_status = $service.Status
	Write-Output "Service Status: $restarted_service_status"
	sleep -Seconds 10
} until ($restarted_service_status -eq "Running" -or $check_service_attempt -eq 0)

if(($restarted_service_status -eq "Stopped") -and ($check_service_attempt -eq 0))
{
	Write-Output "Failed to start service $service_name"
	exit 1
}


#############################################################################################
# Deploy Web API																			#
#############################################################################################
$file_name = "$destination\agent-service\Pasco.Middleware.WebApi.SetParameters.xml"
#$match_rule = '(?<=Server=MTWMID01;Database=)([^"]*)(?=;Trusted_Connection=True;)' #look up the database name in between strings
#Define the correct site name
(Get-Content $file_name) -replace 'Default Web Site/Pasco.Middleware.WebApi_deploy', "$iis_site_name" | Set-Content $file_name #replace the site name
Write-Output "Set Site Name as: $iis_site_name"

#(Get-Content $file_name) -replace $match_rule, $database_name | Set-Content $file_name #replace the connection string
#Write-Output "Set database name as : $database_name. Start deploying Web API..."


$file_path = "$destination\agent-service"

pushd $file_path #set working directory to be the temp folder
	.\Pasco.Middleware.WebApi.deploy.cmd /Y
popd

$api_config_file = Get-Item -Path "$destination\web-api\$config_file"
copy-item $api_config_file -destination "$destination\web-api\Web.config" -force
write-output "Copied $api_config_file as Web.config"

#restart the IIS site to pick up the replaced web.config
Import-Module WebAdministration
Write-output "Stopping $iis_site_name..."
Stop-Website $iis_site_name

$website_state = (Get-WebsiteState -Name $iis_site_name).Value
Write-Output "Web API Status: $website_state"

Write-output "Starting $iis_site_name..."
Start-Website $iis_site_name

$website_state = (Get-WebsiteState -Name $iis_site_name).Value
Write-Output "Web API Status: $website_state"


#############################################################################################
# Clear unwanted files																		#
#############################################################################################
rm -r "$destination\agent-service\configuration"
rm -r "$destination\agent-service\Pasco.Middleware.WebApi*"
rm -r "$destination\web-api\configuration"
write-output "File clean up finished."