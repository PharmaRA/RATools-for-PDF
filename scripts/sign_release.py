"""Optional Windows certificate-store signing. Dry plan unless --execute is explicit.

No private-key files, passwords, certificate enumeration or timestamp services.
An environment selector refers only to a public certificate-store SHA1 thumbprint.
On successful Authenticode verification the bundle manifest is refreshed.
"""

import argparse
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from scripts.release_metadata import signing_plan, record_signing_result, signing_preflight
from ratools_pdf.services.file_safety import check_unlinked_path


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('executable',type=Path)
    parser.add_argument('--certificate-store-sha1')
    parser.add_argument('--certificate-env')
    parser.add_argument('--execute',action='store_true')
    args = parser.parse_args(argv)
    plan = signing_plan(args.certificate_store_sha1,args.certificate_env)
    if not args.execute:
        print(json.dumps(plan))
        return 0
    if os.name != 'nt' or plan['status'] != 'pending':
        parser.error('Execution requires Windows and an explicit certificate-store reference')
    exe = signing_preflight(args.executable,checker=check_unlinked_path)
    selector = '$env:'+args.certificate_env if args.certificate_env else "'"+args.certificate_store_sha1+"'"
    # Values/diagnostics are deliberately never printed. Certificate material
    # stays in the Windows store; only the public selector is resolved there.
    script = "$ErrorActionPreference='Stop'; $selector="+selector+"; if ($selector -notmatch '^[0-9a-fA-F]{40}$') { exit 2 }; "
    script += "$cert=Get-Item -LiteralPath ('Cert:\\CurrentUser\\My\\'+$selector); "
    script += "$result=Set-AuthenticodeSignature -LiteralPath $env:RATOOLS_RELEASE_SIGN_TARGET -Certificate $cert -HashAlgorithm SHA256; "
    script += "if ($result.Status -ne 'Valid') { exit 3 }; $checked=Get-AuthenticodeSignature -LiteralPath $env:RATOOLS_RELEASE_SIGN_TARGET; if ($checked.Status -ne 'Valid') { exit 4 }"
    child_environment = dict(os.environ,RATOOLS_RELEASE_SIGN_TARGET=str(exe))
    success = False
    try:
        result = subprocess.run(['powershell.exe','-NoProfile','-NonInteractive','-Command',script],
            env=child_environment,capture_output=True,creationflags=subprocess.CREATE_NO_WINDOW)
        success = result.returncode == 0
    finally:
        plan = record_signing_result(exe,plan,success)
    print(json.dumps(plan))
    return 0 if success else 1


if __name__ == '__main__':
    raise SystemExit(main())
