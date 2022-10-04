# QueueChecker for WOTLK Classic

#############
# FUNCTIONS #
#############

function Get-ScreenColor {

    [CmdletBinding(DefaultParameterSetName='None')]

    param(
        [Parameter(
            Mandatory=$true,
            ParameterSetName="Pos"
        )]
        [Int]
        $X,
        [Parameter(
            Mandatory=$true,
            ParameterSetName="Pos"
        )]
        [Int]
        $Y
    )
    Add-Type -Assembly system.drawing
    if ($PSCmdlet.ParameterSetName -eq 'None') {
        $pos = [System.Windows.Forms.Cursor]::Position
    } else {
        $pos = New-Object psobject
        $pos | Add-Member -MemberType NoteProperty -Name "X" -Value $X
        $pos | Add-Member -MemberType NoteProperty -Name "Y" -Value $Y
    }
    $map = [System.Drawing.Rectangle]::FromLTRB($pos.X, $pos.Y, $pos.X + 1, $pos.Y + 1)
    $bmp = New-Object System.Drawing.Bitmap(1,1)
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $graphics.CopyFromScreen($map.Location, [System.Drawing.Point]::Empty, $map.Size)
    $pixel = $bmp.GetPixel(0,0)
    $red = $pixel.R
    $green = $pixel.G
    $blue = $pixel.B
    $result = New-Object psobject
    if ($PSCmdlet.ParameterSetName -eq 'None') {
        $result | Add-Member -MemberType NoteProperty -Name "X" -Value $([System.Windows.Forms.Cursor]::Position).X
        $result | Add-Member -MemberType NoteProperty -Name "Y" -Value $([System.Windows.Forms.Cursor]::Position).Y
    }
    $result | Add-Member -MemberType NoteProperty -Name "Red" -Value $red
    $result | Add-Member -MemberType NoteProperty -Name "Green" -Value $green
    $result | Add-Member -MemberType NoteProperty -Name "Blue" -Value $blue
    return $result
}

#get API key from api.ocr.space and edit that out here
function Upload-Screenshot {
PARAM (
    [string]$apiKey,
    [string]$image = "C:\QueueChecker\text.png"
)

    #POWERSHELL OCR API CALL - V2.0, May 30, 2020
    #In this demo we send an image link to the OCR API and download the text result and the searchable PDF

    #Enter your api key here
    $apiUrl = "https://api.ocr.space/parse/image" 

    #Call API with CURL
    $shutUp = curl.exe -X POST $apiurl -H "apikey:$apikey" -F "file=@$image" -F "language=eng" -F "isOverlayRequired=false" -F "iscreatesearchablepdf=true" | ConvertFrom-Json -OutVariable response

    #Done, write OCR'ed text to log file
    $text = $response.ParsedResults.ParsedText
    Write-Host $text

    $response = $response | select SearchablePDFURL

    return $text
}

function Get-ScreenCapture {
    Add-Type -AssemblyName System.Windows.Forms,System.Drawing
    $captureSize = @{Height = 300; Width = 500}
    $screens = [Windows.Forms.Screen]::AllScreens  | Where-Object {$_.Primary -eq "True"}
    $centerPoint = @{FromTop = ($screens.Bounds.Height) / 2; FromLeft = ($screens.Bounds.Width) /2}
    $box = @{}
    $box.Left    = $centerPoint.FromLeft - $captureSize.Width /2
    $box.Top    = $centerPoint.FromTop - $captureSize.Height /2
    $box.Right  = $box.Left + $captureSize.Width
    $box.Bottom = $box.Top + $captureSize.Height

    $bounds   = [Drawing.Rectangle]::FromLTRB($box.Left,$box.Top,$box.Right,$box.Bottom)
    $bmp      = New-Object System.Drawing.Bitmap ([int]$bounds.width), ([int]$bounds.height)
    $graphics = [Drawing.Graphics]::FromImage($bmp)

    $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size)

    if(!(Test-Path "C:\QueueChecker")) {
        New-Item -ItemType Directory C:\QueueChecker -ErrorAction SilentlyContinue
    }
    $bmp.Save("C:\QueueChecker\text.png")

    $graphics.Dispose()
    $bmp.Dispose()
}


#function for testing
function Get-MouseCoordinates {
    Add-Type -AssemblyName System.Windows.Forms
    Start-Sleep 3
    $X = [System.Windows.Forms.Cursor]::Position.X
    $Y = [System.Windows.Forms.Cursor]::Position.Y
    Write-Output "X: $X | Y: $Y"
}

function QueueChecker {
    PARAM(
        $APIKEY
    )
    try {
            Get-ScreenCapture
            $text = Upload-Screenshot -apiKey $APIKEY
            return $text
    }
    catch {
        {1:<#Do this if a terminating exception happens#>}
    }

}
function Send-DiscordWebhook {
    PARAM(
        $WEBHOOK,
        $content
    )
    try {
        $payload = [PSCustomObject]@{

            content = $content
        }
        #Send over payload, converting it to JSON
        Invoke-RestMethod -Uri $WEBHOOK -Body ($payload | ConvertTo-Json -Depth 4) -Method Post -ContentType 'application/json'
    }
    catch {
        {1:<#Do this if a terminating exception happens#>}
    }
}

################
#   UI CODE    #
################

Add-Type -AssemblyName PresentationFramework
add-type -Assembly System.Drawing

[xml]$xaml = @'
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
Title="QueueChecker" WindowStartupLocation="CenterScreen" Height="250" Width="450" Background="#2f80ed">
    <Grid Margin="0,5,0,0">
        <Grid.RowDefinitions> 
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
            <RowDefinition Height="auto"/>
        </Grid.RowDefinitions>
            <Label Grid.Row="0" Name="API_KEY" Content="https://ocr.space/ api key" Foreground="#ffc0cb"/>
            <TextBox Grid.Row="1" Name="API_KEY_INPUT"/>
            <Label Grid.Row="2" Name="WEBHOOK" Content="Discord webhook" Foreground="#ffc0cb"/>
            <TextBox Grid.Row="3" Name="WEBHOOK_INPUT"/>
            <Label Grid.Row="4" Name="Realm" Content="WoW realm" Foreground="#ffc0cb"/>
            <TextBox Grid.Row="5" Name="REALM_INPUT"/>
            <Button Grid.Row="6" Name="QueueChecker" Content="Start QueueChecker" Background="#faffa7" Foreground="#fd625e" Height="40"/>
            <Label Grid.Row="9" Name="Feedback_Label" Content="Current status: INACTIVE" Foreground="#333333"/>
    </Grid>
</Window>
'@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load( $reader )
}
catch {
    Write-Warning $_.Exception
    throw
}

#Create variables based on form control names.
#Variable will be named as 'var_<control name>'

$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    #"trying item $($_.Name)";
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
    } catch {
        throw
   }
}

#This is for dev purposes
#Get-Variable var_*
$window.Add_Loaded({
    $Window.Icon = "https://blz-contentstack-images.akamaized.net/v3/assets/bltf408a0557f4e4998/bltb7c1db49cad77069/60a81b45b078b00d8a909fce/world-of-warcraft.svg?width=168&format=webply&dpr=2&disable=upscale&quality=80"
})

$var_QueueChecker.Add_Click( {
    if($var_API_KEY_INPUT -and $var_WEBHOOK_INPUT) {
        do {
            # sleep 10 minutes default
            $SleepTimer = 600
            $var_Feedback_Label.Content = "Current status: Monitoring Queue";Update-Gui
            # get text from queue
            $text = QueueChecker -APIKEY $var_API_KEY_INPUT.Text
            # set realm
            $realm = $var_REALM_INPUT.Text
            if($text -match "$Realm is Full") {
                Write-Host "Server is full - checking queue position..."

                ## Check queue pos
                $textIndex = $text.Split(" ").IndexOf("queue:")
                $positionInQueue = $text.Split(" ")[$textIndex+1]
                # testing for edge case with double lines returned
                if($positionInQueue.Length -gt 6) {
                    $positionInQueue = $positionInQueue.Split("`r`n")
                }
                Write-Host "Position in queue: $positionInQueue"

                ## Check estimated time left
                $textIndex = $text.Split(" ").IndexOf("time:")
                $estimatedTime = $text.Split(" ")[$textIndex+1]
                Write-Host "Time left: $estimatedTime min"
                $time = (Get-Date).AddMinutes($estimatedTime).ToString("HH:mm")
                $logFile = "C:\QueueChecker\log.txt"

                # if less than 15 minutes, webhook
                if($estimatedTime -le "15") {
                    $content = "**QueueChecker** -- @everyone -- $estimatedTime remaining until login ready on $realm (around $time)"
                    # sleep 59 seconds after 15 minutes
                    $SleepTimer = 59
                    # if less than 1 min, queueover
                    if($estimatedTime -le "1") {
                        $QueueOver = $true
                    }
                } else {
                    $content = "**QueueChecker** -- $estimatedTime minutes remaining until login ready on $realm (around $time)"
                }
                # send info to the discord
                Send-DiscordWebhook -WEBHOOK $var_WEBHOOK_INPUT.Text -content $content

                # send info to logfile
                Add-Content $logfile $content

                # if queue over, don't sleep
                if(!$QueueOver) {
                    Start-Sleep $SleepTimer
                } else {
                    $var_Feedback_Label.Content = "Current status: Queue Over";Update-Gui
                }
            }
        } until ($QueueOver)
    }
})

function Update-Gui {
    # Basically WinForms Application.DoEvents()
    $Window.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Background, [action]{})
}
$Null = $window.ShowDialog()
