import os
from pathlib import Path
root=Path(__file__).resolve().parents[1]; gradle=root/'android/app/build.gradle'
ks=os.environ['SFC_KEYSTORE_PATH']; alias=os.environ['SFC_KEY_ALIAS']; sp=os.environ['SFC_KEYSTORE_PASSWORD']; kp=os.environ.get('SFC_KEY_PASSWORD') or sp
s=gradle.read_text(); esc=lambda x:str(x).replace('\\','\\\\').replace("'","\\'")
block=f"\n    signingConfigs {{ release {{ storeFile file('{esc(ks)}'); storePassword '{esc(sp)}'; keyAlias '{esc(alias)}'; keyPassword '{esc(kp)}' }} }}\n"
if 'signingConfigs {' not in s:s=s.replace('android {','android {'+block,1)
if 'signingConfig signingConfigs.release' not in s:
 s=s.replace('buildTypes {','buildTypes {\n        release { signingConfig signingConfigs.release }',1)
gradle.write_text(s)