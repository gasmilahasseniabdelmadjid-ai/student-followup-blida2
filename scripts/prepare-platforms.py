from pathlib import Path
import re
root=Path(__file__).resolve().parents[1]
m=root/'android/app/src/main/AndroidManifest.xml'
if m.exists():
 s=m.read_text()
 if 'android.permission.NFC' not in s:s=s.replace('<manifest xmlns:android="http://schemas.android.com/apk/res/android">','<manifest xmlns:android="http://schemas.android.com/apk/res/android">\n<uses-permission android:name="android.permission.NFC" />',1)
 if 'android.hardware.nfc' not in s:s=s.replace('<application','<uses-feature android:name="android.hardware.nfc" android:required="false" />\n<application',1)
 m.write_text(s)
v=root/'android/variables.gradle'
if v.exists():
 s=v.read_text()
 for k,val in [('compileSdkVersion','36'),('targetSdkVersion','36'),('minSdkVersion','24')]:s=re.sub(rf'{k}\s*=\s*\d+',f'{k} = {val}',s)
 v.write_text(s)
print('Platform preparation completed')