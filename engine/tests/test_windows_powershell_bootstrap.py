"""Verify ZIP bootstrap in a new, profile-free Windows PowerShell 5.1 process.

No COM, source PPT, or user output is touched. Only the production initializer
is extracted by the PowerShell AST; the rest of ppt_render.ps1 is never run.
Run directly with --report PATH to retain the fresh-process evidence as JSON.
"""
from __future__ import annotations

import base64
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
WINDOWS_POWERSHELL = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32/WindowsPowerShell/v1.0/powershell.exe"
LAST_REPORT = {}


PROBE = r'''
$ErrorActionPreference = 'Stop'
$initialAssemblies = @([AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetName().Name })
$initiallyResolved = $null -ne ('System.IO.Compression.ZipArchiveMode' -as [type])
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($env:FIGURE_TEST_RENDERER, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -ne 0) { throw ('Renderer parse failed: ' + ($parseErrors.Message -join '; ')) }
$initializer = $ast.Find({ param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Initialize-ZipSupport'
}, $true)
if ($null -eq $initializer) { throw 'Production Initialize-ZipSupport function was not found.' }
# Define and invoke precisely the production function. This never runs COM or
# any top-level renderer statement and never dot-sources a presentation script.
Invoke-Expression $initializer.Extent.Text
$initializerOutput = @(Initialize-ZipSupport)
if ($null -eq ('System.IO.Compression.ZipArchiveMode' -as [type])) { throw 'ZipArchiveMode remains unavailable after bootstrap.' }
if ($null -eq ('System.IO.Compression.ZipFile' -as [type])) { throw 'ZipFile remains unavailable after bootstrap.' }
$zipPath = [IO.Path]::Combine($env:FIGURE_TEST_DIRECTORY, 'fresh-process.zip')
$archive = [IO.Compression.ZipFile]::Open($zipPath, [IO.Compression.ZipArchiveMode]::Create)
try {
    $entry = $archive.CreateEntry('original.txt')
    $writer = New-Object IO.StreamWriter($entry.Open())
    try { $writer.Write('original content') } finally { $writer.Dispose() }
} finally { $archive.Dispose() }
$archive = [IO.Compression.ZipFile]::Open($zipPath, [IO.Compression.ZipArchiveMode]::Update)
try {
    $entry = $archive.GetEntry('original.txt')
    if ($null -eq $entry) { throw 'Created entry is missing.' }
    $entry.Delete()
    $entry = $archive.CreateEntry('updated.txt')
    $writer = New-Object IO.StreamWriter($entry.Open())
    try { $writer.Write('updated content') } finally { $writer.Dispose() }
} finally { $archive.Dispose() }
$archive = [IO.Compression.ZipFile]::Open($zipPath, [IO.Compression.ZipArchiveMode]::Read)
try {
    if ($null -ne $archive.GetEntry('original.txt')) { throw 'Deleted entry remains.' }
    $entry = $archive.GetEntry('updated.txt')
    if ($null -eq $entry) { throw 'Updated entry is missing.' }
    $reader = New-Object IO.StreamReader($entry.Open())
    try { $contents = $reader.ReadToEnd() } finally { $reader.Dispose() }
    if ($contents -ne 'updated content') { throw 'ZIP update/read content mismatch.' }
} finally { $archive.Dispose() }
[ordered]@{
    status = 'pass'
    powershell_version = $PSVersionTable.PSVersion.ToString()
    powershell_edition = $PSVersionTable.PSEdition
    process_path = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    no_profile = $true
    process_id = $PID
    initially_loaded_compression_assembly = $initialAssemblies -contains 'System.IO.Compression'
    zip_archive_mode_initially_resolved = $initiallyResolved
    zip_archive_mode_after_bootstrap = ('System.IO.Compression.ZipArchiveMode' -as [type]).FullName
    zip_file_after_bootstrap = ('System.IO.Compression.ZipFile' -as [type]).FullName
    compression_assembly = [IO.Compression.ZipArchiveMode].Assembly.FullName
    file_system_assembly = [IO.Compression.ZipFile].Assembly.FullName
    operations = @('create', 'update', 'delete entry', 'read')
    com_used = $false
    initializer_output = $initializerOutput
} | ConvertTo-Json -Depth 6 -Compress
'''


def run_probe():
    """Start a fresh Windows PowerShell process; do not inherit loaded modules."""
    env = os.environ.copy()
    # A PowerShell 7 parent may export its private module paths. Reconstructing
    # the Windows defaults is the child shell's job, rather than this test's.
    env.pop('PSModulePath', None)
    env['FIGURE_TEST_RENDERER'] = str(ROOT / 'ppt_render.ps1')
    with tempfile.TemporaryDirectory(prefix='figure-zip-bootstrap-') as tmp:
        env['FIGURE_TEST_DIRECTORY'] = tmp
        encoded = base64.b64encode(PROBE.encode('utf-16le')).decode('ascii')
        completed = subprocess.run(
            [str(WINDOWS_POWERSHELL), '-NoProfile', '-NonInteractive',
             '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encoded],
            env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            timeout=60, check=False,
        )
        # The report is ASCII JSON even on a system whose shell output code page
        # differs from Python's default locale.
        stdout = completed.stdout.decode('utf-8-sig', errors='replace').strip()
        stderr = completed.stderr.decode('utf-8-sig', errors='replace').strip()
        if completed.returncode:
            raise AssertionError(f'Fresh PowerShell exit {completed.returncode}:\n{stdout}\n{stderr}')
        return json.loads(stdout.splitlines()[-1])


@unittest.skipUnless(os.name == 'nt' and WINDOWS_POWERSHELL.exists(), 'Requires Windows PowerShell 5.1')
class FreshPowerShellBootstrapTests(unittest.TestCase):
    def test_fresh_windows_powershell_zip_create_update_read(self):
        report = run_probe()
        LAST_REPORT.update(report)
        self.assertEqual(report['status'], 'pass')
        self.assertTrue(report['powershell_version'].startswith('5.1.'))
        self.assertEqual(report['powershell_edition'], 'Desktop')
        self.assertFalse(report['com_used'])
        self.assertEqual(report['operations'], ['create', 'update', 'delete entry', 'read'])
        self.assertEqual(report['zip_archive_mode_after_bootstrap'], 'System.IO.Compression.ZipArchiveMode')
        self.assertEqual(report['zip_file_after_bootstrap'], 'System.IO.Compression.ZipFile')
        # Normally false in clean Windows PowerShell. Record it rather than
        # assume host-specific preload policy; the child is always newly created.
        self.assertIn('zip_archive_mode_initially_resolved', report)


if __name__ == '__main__':
    report_path = None
    argv = list(sys.argv)
    if '--report' in argv:
        index = argv.index('--report')
        report_path = Path(argv[index + 1])
        del argv[index:index + 2]
    suite = unittest.main(argv=argv, verbosity=2, exit=False)
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(LAST_REPORT, indent=2), encoding='utf-8')
    sys.exit(0 if suite.result.wasSuccessful() else 1)
