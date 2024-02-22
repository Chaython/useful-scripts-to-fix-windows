# Import the AudioDeviceCmdlets module
Import-Module AudioDeviceCmdlets

# Define the indices of the target devices 
$deviceIndex1 = 1
$deviceIndex2 = 2

# Function to toggle between the desired indices
function Toggle-AudioDevices {
    $targetIndex = if ((Get-AudioDevice -Index $deviceIndex1).ID -eq (Get-AudioDevice -Playback).ID) {
        $deviceIndex2 
    } else {
        $deviceIndex1
    }

    Set-AudioDevice -Index $targetIndex
    Set-AudioDevice -Communication -Index $targetIndex

    Write-Host "Default playback and communication devices switched to index $targetIndex"
}

# Toggle playback and communication devices
Toggle-AudioDevices 