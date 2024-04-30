param(
    [Parameter(ValueFromPipeline=$true)]
    [string]$InputString
)

begin {
    $sb = New-Object System.Text.StringBuilder
}

process {
    # Use the input string from the pipeline if available, otherwise use the Base64EncodedString parameter
    if ($InputString.length -eq 0)
    {
        break
    }

    # Decode the Base64 encoded string
    $decodedBytes = [System.Convert]::FromBase64String($InputString)

    # Gzip decompress the decoded bytes
    $gzipStream = New-Object System.IO.MemoryStream
    $gzipStream.Write($decodedBytes, 0, $decodedBytes.Length)
    [void]$gzipStream.Seek(0, [System.IO.SeekOrigin]::Begin)
    $gzipStream.Flush()
    $gzipStream.Position = 0
    $gzipStream = New-Object System.IO.Compression.GzipStream($gzipStream, [System.IO.Compression.CompressionMode]::Decompress)
    $reader = New-Object System.IO.StreamReader($gzipStream)
    
    # Iterate over the bytes and write them to the output as ASCII
    $bufferSize = 1024
    $buffer = New-Object byte[] $bufferSize
    $bytesRead = $reader.BaseStream.Read($buffer, 0, $bufferSize)
    $hexString = [string]::Join(",", $($buffer[0..($bytesRead-1)] | ForEach-Object { $_.ToString("x2") }))
    [void]$sb.Append($hexString)
    $bytesRead = $reader.BaseStream.Read($buffer, 0, $bufferSize)
    while ($bytesRead -gt 0) {
        $hexString = [string]::Join(",", $($buffer[0..($bytesRead-1)] | ForEach-Object { $_.ToString("x2") }))
        [void]$sb.Append("," + $hexString)
        $bytesRead = $reader.BaseStream.Read($buffer, 0, $bufferSize)
    }

    $reader.Close()
    $gzipStream.Close()
}

end {
    # Output the ASCII representation of the bytes
    $sb.ToString()
}