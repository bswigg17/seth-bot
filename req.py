import requests
import os
current_version: str = "1.0"


version = requests.get("https://raw.githubusercontent.com/bswigg17/elera-tool-updates/master/version.txt").text

if (float(version) > float(current_version)):
    res = requests.get("https://raw.githubusercontent.com/bswigg17/elera-tool-updates/master/new")
    if (res.status_code == 200):
        with open('./dist/new', 'wb') as f:
            for chunk in res.iter_content(chunk_size=1024):
                if chunk:
                    f.write(chunk)
        os.system("chmod u=rwx,g=r,o=r ./dist/new")
        os.system("./dist/new")
        current_version = version

