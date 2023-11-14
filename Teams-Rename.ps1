# Use this to get around the new M$ Teams needing a Premium SKU to have backgrounds for meetings.
# This will Take any image files you have, rename to the new naming convention, create a _thumb file (and resize to the correct dimentions) and move it to your desired location.
# This can be a server location, if you have the required permission to the share you have selected.
# Images shoult be stored here %LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\Uploads
# 
# Import the required .NET assemblies for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to create GUI
function Create-GUI {
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Teams Background Name Change"
    $form.Size = New-Object System.Drawing.Size(400, 170)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedSingle"  # Set the form border style to fixed single
    $form.MinimumSize = $form.Size  # Set the minimum size to prevent resizing
    $form.MaximumSize = $form.Size  # Set the maximum size to prevent resizing

    # Create labels
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10, 20)
    $label1.Size = New-Object System.Drawing.Size(200, 20)
    $label1.Text = "Source Folder Path:"
    $form.Controls.Add($label1)

    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10, 50)
    $label2.Size = New-Object System.Drawing.Size(200, 20)
    $label2.Text = "Destination Folder Path:"
    $form.Controls.Add($label2)

    # Create textboxes
    $textbox1 = New-Object System.Windows.Forms.TextBox
    $textbox1.Location = New-Object System.Drawing.Point(220, 20)
    $textbox1.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($textbox1)

    $textbox2 = New-Object System.Windows.Forms.TextBox
    $textbox2.Location = New-Object System.Drawing.Point(220, 50)
    $textbox2.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($textbox2)

    # Create buttons
    $button1 = New-Object System.Windows.Forms.Button
    $button1.Location = New-Object System.Drawing.Point(30, 80)
    $button1.Size = New-Object System.Drawing.Size(150, 30)
    $button1.Text = "Rename and Copy"
    $button1.Add_Click({
        Rename-Copy-Images $textbox1.Text $textbox2.Text
    })
    $form.Controls.Add($button1)

    $button2 = New-Object System.Windows.Forms.Button
    $button2.Location = New-Object System.Drawing.Point(200, 80)
    $button2.Size = New-Object System.Drawing.Size(150, 30)
    $button2.Text = "Close"
    $button2.Add_Click({ $form.Close() })
    $form.Controls.Add($button2)

    # Display the form
    $form.ShowDialog()
}

# Function to rename, copy, and create thumbnails of images with GUID-style filename (new Teams standard)
function Rename-Copy-Images($sourcePath, $destinationPath) {
    try {
        # Check if source folder exists
        if (Test-Path $sourcePath) {
            # Check if destination folder exists, if not, create it
            if (!(Test-Path $destinationPath)) {
                New-Item -ItemType Directory -Path $destinationPath | Out-Null
            }

            # Get all files with specified image extensions in the source folder
            $imageFiles = Get-ChildItem $sourcePath -File | Where-Object { $_.Extension -in ('.jpg', '.png', '.gif') }

            if ($imageFiles.Count -eq 0) {
                Write-Host 'No File Found in Source Location' -ForegroundColor Yellow
            }
            else {
                foreach ($file in $imageFiles) {
                    try {
                        $guid = [System.Guid]::NewGuid()
                        $newName = "$guid$($file.Extension)"

                        $newPath = Join-Path $destinationPath $newName
                        Copy-Item $file.FullName -Destination $newPath

                        # Create thumbnail filename by appending "_thumb" before the extension
                        $thumbName = "$guid" + "_thumb" + "$($file.Extension)"
                        $thumbPath = Join-Path $destinationPath $thumbName

                        if ($thumbName -like '*_thumb*') {
                            # Resize and create a thumbnail only for files containing "_thumb" in the filename
                            Add-Type -AssemblyName System.Drawing
                            $image = [System.Drawing.Image]::FromFile($newPath)
                            $thumbnail = $image.GetThumbnailImage(280, 158, $null, [System.IntPtr]::Zero)
                            $thumbnail.Save($thumbPath, $image.RawFormat)
                        }
                    }
                    catch {
                        # Handle exceptions if necessary
                    }
                }

                Write-Host 'Process Completed. Files Renamed, Copied and Thumbnails Created.' -ForegroundColor Green
            }
        } else {
            Write-Host 'Source folder not found.' -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host 'Process Completed.' -ForegroundColor Green
    }
}

# Create the GUI
Create-GUI
