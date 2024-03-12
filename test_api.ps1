# Define the path to your PowerPoint file
$fileToUpload = "file.pptx"

# Define the URL of your Flask API endpoint
$uri = "http://localhost:8080/shrink_pptx"

# Create a multipart/form-data content type header
$multipartContent = [System.Net.Http.MultipartFormDataContent]::new()

# Add the PowerPoint file to the multipart/form-data content
$fileStream = [System.IO.FileStream]::new($fileToUpload, [System.IO.FileMode]::Open)
$fileHeader = [System.Net.Http.Headers.ContentDispositionHeaderValue]::new("form-data")
$fileHeader.Name = "file"
$fileHeader.FileName = $fileToUpload
$content = [System.Net.Http.StreamContent]::new($fileStream)
$content.Headers.ContentDisposition = $fileHeader
$multipartContent.Add($content)

# Send the request
$response = Invoke-WebRequest -Uri $uri -Method Post -Body $multipartContent -ContentType "multipart/form-data"

# Close the file stream
$fileStream.Close()

# Display the response status and headers
Write-Output "Response status: $($response.StatusCode)"
Write-Output "Response headers: $($response.Headers)"

# Optionally, save the response to a file
$shrunkFilePath = "shrunk_file.pptx"
[System.IO.File]::WriteAllBytes($shrunkFilePath, $response.Content)
