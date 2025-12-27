# Word/PowerPoint to PDF/Image Converter Script
# Uses Microsoft Word/PowerPoint COM automation for 1:1 conversion

param(
    [Parameter(Mandatory=$true)]
    [string]$InputPath,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$true)]
    [ValidateSet("pdf", "image")]
    [string]$Format
)

$ErrorActionPreference = "Stop"

# Determine file type
$extension = [System.IO.Path]::GetExtension($InputPath).ToLower()
$isPowerPoint = ($extension -eq ".pptx") -or ($extension -eq ".ppt")

try {
    if ($isPowerPoint) {
        # PowerPoint handling
        $ppt = New-Object -ComObject PowerPoint.Application
        # $ppt.Visible = $false  # PowerPoint requires visibility in some versions
        
        # Open presentation
        $presentation = $ppt.Presentations.Open($InputPath, $true, $false, $false) # ReadOnly, Untitled, WithWindow
        
        if ($Format -eq "pdf") {
            # Export as PDF
            # ppSaveAsPDF = 32
            $presentation.SaveAs($OutputPath, 32)
            Write-Output "PDF_SUCCESS:$OutputPath"
        }
        elseif ($Format -eq "image") {
            # Create temp PDF first
            $tempPdf = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "temp_ppt_export.pdf")
            $presentation.SaveAs($tempPdf, 32)
            
            Write-Output "IMAGE_TEMP_PDF:$tempPdf"
            Write-Output "IMAGE_OUTPUT:$OutputPath"
        }
        
        # Cleanup
        $presentation.Close()
        $ppt.Quit()
        
        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
    }
    else {
        # Word handling
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0  # wdAlertsNone
        
        # Open document
        $doc = $word.Documents.Open($InputPath)
        
        if ($Format -eq "pdf") {
            # Export as PDF
            # wdExportFormatPDF = 17
            # wdExportOptimizeForPrint = 0
            $doc.ExportAsFixedFormat(
                $OutputPath,
                17,  # wdExportFormatPDF
                $false,  # OpenAfterExport
                0,  # wdExportOptimizeForPrint
                0,  # Range - wdExportAllDocument
                1,  # From
                1,  # To
                7,  # Item - wdExportDocumentWithMarkup
                $true,  # IncludeDocProps
                $true,  # KeepIRM
                0,  # CreateBookmarks - wdExportCreateNoBookmarks
                $true,  # DocStructureTags
                $true,  # BitmapMissingFonts
                $false  # UseISO19005_1 (PDF/A)
            )
            Write-Output "PDF_SUCCESS:$OutputPath"
        }
        elseif ($Format -eq "image") {
            # Create temp PDF first
            $tempPdf = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "temp_word_export.pdf")
            $doc.ExportAsFixedFormat($tempPdf, 17, $false, 0, 0, 1, 1, 7, $true, $true, 0, $true, $true, $false)
            
            Write-Output "IMAGE_TEMP_PDF:$tempPdf"
            Write-Output "IMAGE_OUTPUT:$OutputPath"
        }
        
        # Cleanup
        $doc.Close([ref]$false)
        $word.Quit()
        
        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    exit 0
}
catch {
    Write-Error "ERROR:$($_.Exception.Message)"
    
    # Cleanup on error
    if ($isPowerPoint) {
        if ($presentation) { 
            try { $presentation.Close() } catch {}
        }
        if ($ppt) { 
            try { $ppt.Quit() } catch {}
        }
    }
    else {
        if ($doc) { 
            try { $doc.Close([ref]$false) } catch {}
        }
        if ($word) { 
            try { $word.Quit() } catch {}
        }
    }
    
    exit 1
}
