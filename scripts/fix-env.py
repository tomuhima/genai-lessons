#!/usr/bin/env python3
import re

with open('/root/n8n-compose/.env', 'r') as f:
    content = f.read()

key_match = re.search(r'CLAUDE_API_KEY=([^\n\r\[]+)', content)
claude_key = key_match.group(1).strip() if key_match else ''

email = 'mahyy8810' + chr(64) + 'gmail' + chr(46) + 'com'
domain = 'xvps' + chr(46) + 'jp'

new_content = '\n'.join([
    'SSL_EMAIL=' + email,
    'SUBDOMAIN=n8n-light-m-n',
    'DOMAIN_NAME=' + domain,
    'GENERIC_TIMEZONE=Asia/Tokyo',
    'CLAUDE_API_KEY=' + claude_key,
]) + '\n'

with open('/root/n8n-compose/.env', 'w') as f:
    f.write(new_content)

print('Done:')
print(new_content)
