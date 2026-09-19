#!/bin/bash
# クラウド環境でChromium(Playwright)がHTTPSサイトを開けるようにする。
# 原因：プロキシのCA証明書がNSSストアに未登録で ERR_CERT_AUTHORITY_INVALID になる。
# 使い方：bash _build/setup_browser_ca.sh  （セッションごとに1回。数十秒）
set -e

if certutil -L -d sql:$HOME/.pki/nssdb 2>/dev/null | grep -q agentproxy; then
  echo "登録済み。何もしません。"
  exit 0
fi

command -v certutil >/dev/null || { apt-get update -q && apt-get install -y libnss3-tools -q; }
mkdir -p $HOME/.pki/nssdb
certutil -d sql:$HOME/.pki/nssdb -L >/dev/null 2>&1 || certutil -N -d sql:$HOME/.pki/nssdb --empty-password

python3 - <<'EOF'
import re
data = open('/root/.ccr/ca-bundle.crt').read()
certs = re.findall(r'-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----', data, re.S)
for i, c in enumerate(certs):
    open(f'/tmp/agentproxy_ca_{i}.pem', 'w').write(c)
print(f'{len(certs)} certs extracted')
EOF

for f in /tmp/agentproxy_ca_*.pem; do
  certutil -A -d sql:$HOME/.pki/nssdb -n "agentproxy-$(basename $f .pem)" -t "C,," -i "$f"
done
rm -f /tmp/agentproxy_ca_*.pem
echo "完了。Chromium(executable_path=/opt/pw-browsers/chromium)でHTTPSが開けます。"
