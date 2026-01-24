Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- DEFINE WINDOWS API TO CHECK MOUSE BUTTONS ---
$signature = @"
[DllImport("user32.dll")]
public static extern short GetAsyncKeyState(int vKey);
"@
$User32 = Add-Type -MemberDefinition $signature -Name "User32" -Namespace Win32 -PassThru

# --- CONFIGURATION ---
$ResistanceTimeMs = 200   # Time (ms) to push against edge before wrapping
$PollRateMs       = 10    # Refresh rate
# ---------------------

function Start-CursorWrapper {
    Write-Host "Cursor Wrapping Active." -ForegroundColor Cyan
    Write-Host "  - Resistance: $ResistanceTimeMs ms" -ForegroundColor Gray
    Write-Host "  - Drag Protection: ON (Holding Left Click blocks wrap)" -ForegroundColor Green
    Write-Host "  - Press CTRL + C to stop." -ForegroundColor Yellow
    Write-Host "--------------------------------------"

    $edgeTimer = [System.Diagnostics.Stopwatch]::New()
    $edgeActive = $false
    $lastEdge = ""

    # Virtual Key Code for Left Mouse Button is 0x01
    $VK_LBUTTON = 0x01

    try {
        while ($true) {
            $cursorPos = [System.Windows.Forms.Cursor]::Position
            $screen = [System.Windows.Forms.Screen]::GetBounds($cursorPos)
            
            # Define Edges
            $leftEdge   = $screen.X
            $rightEdge  = $screen.X + $screen.Width - 1
            $topEdge    = $screen.Y
            $bottomEdge = $screen.Y + $screen.Height - 1

            # Check if Left Mouse Button is currently held down
            # GetAsyncKeyState returns a short; if the Most Significant Bit (0x8000) is set, the key is down.
            $isDragging = ($User32::GetAsyncKeyState($VK_LBUTTON) -band 0x8000) -ne 0

            # Identify if we are at an edge
            $atEdge = $false
            $currentEdge = ""

            if ($cursorPos.X -le $leftEdge)       { $atEdge = $true; $currentEdge = "Left" }
            elseif ($cursorPos.X -ge $rightEdge)  { $atEdge = $true; $currentEdge = "Right" }
            elseif ($cursorPos.Y -le $topEdge)    { $atEdge = $true; $currentEdge = "Top" }
            elseif ($cursorPos.Y -ge $bottomEdge) { $atEdge = $true; $currentEdge = "Bottom" }

            # --- WRAPPING LOGIC ---
            # We only wrap if:
            # 1. We are at an edge
            # 2. We are NOT dragging ($isDragging is false)
            if ($atEdge -and -not $isDragging) {
                
                # Start timer if this is a new edge contact
                if (-not $edgeActive -or $currentEdge -ne $lastEdge) {
                    $edgeTimer.Restart()
                    $edgeActive = $true
                    $lastEdge = $currentEdge
                }

                # If resistance time passed, wrap
                if ($edgeTimer.ElapsedMilliseconds -ge $ResistanceTimeMs) {
                    $newX = $cursorPos.X; $newY = $cursorPos.Y

                    switch ($currentEdge) {
                        "Left"   { $newX = $rightEdge - 2 }
                        "Right"  { $newX = $leftEdge + 2 }
                        "Top"    { $newY = $bottomEdge - 2 }
                        "Bottom" { $newY = $topEdge + 2 }
                    }

                    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($newX, $newY)
                    
                    # Reset
                    $edgeTimer.Reset()
                    $edgeActive = $false
                    Start-Sleep -Milliseconds 100
                }
            }
            else {
                # Reset if not at edge OR if dragging
                $edgeTimer.Reset()
                $edgeActive = $false
                $lastEdge = ""
            }

            Start-Sleep -Milliseconds $PollRateMs
        }
    }
    finally {
        Write-Host "`nStopping Cursor Wrapper." -ForegroundColor Red
    }
}

Start-CursorWrapper