# Re-export Anderson_BM_NSF_Synergistic_Activities.pdf from the .docx with Word,
# then confirm with pypdf that it is exactly one page.
#
#   powershell -ExecutionPolicy Bypass -File export_pdf.ps1
#
# If Word reports a file-lock error, close the .docx in Word first.

$dir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$docx = Join-Path $dir "Anderson_BM_NSF_Synergistic_Activities.docx"
$pdf  = Join-Path $dir "Anderson_BM_NSF_Synergistic_Activities.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($docx)
    $pages = $doc.ComputeStatistics(2)          # wdStatisticPages
    $rng = $doc.Content
    $rng.Collapse(0)                            # wdCollapseEnd
    $y = $rng.Information(6)                    # wdVerticalPositionRelativeToPage
    Write-Output ("Word: pages={0}; top of last line at {1:N1} pt (text area ends at 720 pt)" -f $pages, $y)
    $doc.SaveAs([ref]$pdf, [ref]17)             # wdFormatPDF
    $doc.Close($false)
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

python -I -c "import sys; from pypdf import PdfReader; n=len(PdfReader(sys.argv[1]).pages); print('pypdf pages:', n); sys.exit(0 if n == 1 else 1)" $pdf
if (-not $?) { Write-Error "PDF is not one page: tighten the prose in build_synergistic_activities.py (do not shrink margins or font)." }
