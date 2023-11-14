from api import WUApi
from utils import b64hdec, parse_xml, parse_xml_html
from tempfile import TemporaryDirectory
from subprocess import run, Popen, PIPE
from os import chdir, getcwd, makedirs, remove, environ, listdir
from os.path import join, exists, abspath, isdir
from shutil import copyfile, move, rmtree
from bs4 import BeautifulSoup
from datetime import datetime
from collections import defaultdict
from json import dumps
import re
import argparse

LATEST_RETAIL = "19044"
VERSION_BUILD = {
    "20h2": ("19042", "vb_release"),
    "21h1": ("19043", "vb_release"),
    "21h2": ("19044", "vb_release"),
    "22h1": ("22000", "ni_release"),
    "22h2": ("22621", "ni_release")
}
UUP_DIR = "UUP"
TEMP_DIR = "TEMP"
ISO_DIR = "ISODIR"
BOOT_SRC = [
	"sources/alert.gif",
	"sources/api-ms-win-core-apiquery-l1-1-0.dll",
	"sources/api-ms-win-downlevel-advapi32-l1-1-0.dll",
	"sources/api-ms-win-downlevel-advapi32-l1-1-1.dll",
	"sources/api-ms-win-downlevel-advapi32-l2-1-0.dll",
	"sources/api-ms-win-downlevel-advapi32-l2-1-1.dll",
	"sources/api-ms-win-downlevel-advapi32-l3-1-0.dll",
	"sources/api-ms-win-downlevel-advapi32-l4-1-0.dll",
	"sources/api-ms-win-downlevel-kernel32-l1-1-0.dll",
	"sources/api-ms-win-downlevel-kernel32-l2-1-0.dll",
	"sources/api-ms-win-downlevel-ole32-l1-1-0.dll",
	"sources/api-ms-win-downlevel-ole32-l1-1-1.dll",
	"sources/api-ms-win-downlevel-shlwapi-l1-1-0.dll",
	"sources/api-ms-win-downlevel-shlwapi-l1-1-1.dll",
	"sources/api-ms-win-downlevel-user32-l1-1-0.dll",
	"sources/api-ms-win-downlevel-user32-l1-1-1.dll",
	"sources/api-ms-win-downlevel-version-l1-1-0.dll",
	"sources/appcompat.xsl",
	"sources/appcompat_bidi.xsl",
	"sources/appcompat_detailed_bidi_txt.xsl",
	"sources/appcompat_detailed_txt.xsl",
	"sources/appraiser.dll",
	"sources/ARUNIMG.dll",
	"sources/arunres.dll",
	"sources/autorun.dll",
	"sources/bcd.dll",
	"sources/bootsvc.dll",
	"sources/cmisetup.dll",
	"sources/compatctrl.dll",
	"sources/compatprovider.dll",
	"sources/compliance.ini",
	"sources/cryptosetup.dll",
	"sources/diager.dll",
	"sources/diagnostic.dll",
	"sources/diagtrack.dll",
	"sources/diagtrackrunner.exe",
	"sources/dism.exe",
	"sources/dismapi.dll",
	"sources/dismcore.dll",
	"sources/dismcoreps.dll",
	"sources/dismprov.dll",
	"sources/ext-ms-win-advapi32-encryptedfile-l1-1-0.dll",
	"sources/folderprovider.dll",
	"sources/hwcompat.dll",
	"sources/hwcompat.txt",
	"sources/hwexclude.txt",
	"sources/idwbinfo.txt",
	"sources/imagelib.dll",
	"sources/imagingprovider.dll",
	"sources/input.dll",
	"sources/lang.ini",
	"sources/locale.nls",
	"sources/logprovider.dll",
	"sources/MediaSetupUIMgr.dll",
	"sources/ndiscompl.dll",
	"sources/nlsbres.dll",
	"sources/ntdsupg.dll",
	"sources/offline.xml",
	"sources/pnpibs.dll",
	"sources/reagent.admx",
	"sources/reagent.dll",
	"sources/reagent.xml",
	"sources/rollback.exe",
	"sources/schema.dat",
	"sources/segoeui.ttf",
	"sources/ServicingCommon.dll",
	"sources/setup.exe",
	"sources/setupcompat.dll",
	"sources/SetupCore.dll",
	"sources/SetupHost.exe",
	"sources/SetupMgr.dll",
	"sources/SetupPlatform.cfg",
	"sources/SetupPlatform.dll",
	"sources/SetupPlatform.exe",
	"sources/SetupPrep.exe",
	"sources/SmiEngine.dll",
	"sources/spflvrnt.dll",
	"sources/spprgrss.dll",
	"sources/spwizeng.dll",
	"sources/spwizimg.dll",
	"sources/spwizres.dll",
	"sources/sqmapi.dll",
	"sources/testplugin.dll",
	"sources/unattend.dll",
	"sources/unbcl.dll",
	"sources/upgloader.dll",
	"sources/upgrade_frmwrk.xml",
	"sources/utcapi.dll",
	"sources/uxlib.dll",
	"sources/uxlibres.dll",
	"sources/vhdprovider.dll",
	"sources/w32uiimg.dll",
	"sources/w32uires.dll",
	"sources/warning.gif",
	"sources/wdsclient.dll",
	"sources/wdsclientapi.dll",
	"sources/wdscommonlib.dll",
	"sources/wdscore.dll",
	"sources/wdscsl.dll",
	"sources/wdsimage.dll",
	"sources/wdstptc.dll",
	"sources/wdsutil.dll",
	"sources/wimgapi.dll",
	"sources/wimprovider.dll",
	"sources/win32ui.dll",
	"sources/WinDlp.dll",
	"sources/winsetup.dll",
	"sources/wpx.dll",
	"sources/xmllite.dll",
	"sources/{mui_lang}/appraiser.dll.mui",
	"sources/{mui_lang}/arunres.dll.mui",
	"sources/{mui_lang}/cmisetup.dll.mui",
	"sources/{mui_lang}/compatctrl.dll.mui",
	"sources/{mui_lang}/compatprovider.dll.mui",
	"sources/{mui_lang}/dism.exe.mui",
	"sources/{mui_lang}/dismapi.dll.mui",
	"sources/{mui_lang}/dismcore.dll.mui",
	"sources/{mui_lang}/dismprov.dll.mui",
	"sources/{mui_lang}/folderprovider.dll.mui",
	"sources/{mui_lang}/imagingprovider.dll.mui",
	"sources/{mui_lang}/input.dll.mui",
	"sources/{mui_lang}/logprovider.dll.mui",
	"sources/{mui_lang}/MediaSetupUIMgr.dll.mui",
	"sources/{mui_lang}/nlsbres.dll.mui",
	"sources/{mui_lang}/pnpibs.dll.mui",
	"sources/{mui_lang}/reagent.adml",
	"sources/{mui_lang}/reagent.dll.mui",
	"sources/{mui_lang}/rollback.exe.mui",
	"sources/{mui_lang}/setup.exe.mui",
	"sources/{mui_lang}/setup_help_upgrade_or_custom.rtf",
	"sources/{mui_lang}/setupcompat.dll.mui",
	"sources/{mui_lang}/SetupCore.dll.mui",
	"sources/{mui_lang}/SetupMgr.dll.mui",
	"sources/{mui_lang}/setupplatform.exe.mui",
	"sources/{mui_lang}/SetupPrep.exe.mui",
	"sources/{mui_lang}/smiengine.dll.mui",
	"sources/{mui_lang}/spwizres.dll.mui",
	"sources/{mui_lang}/upgloader.dll.mui",
	"sources/{mui_lang}/uxlibres.dll.mui",
	"sources/{mui_lang}/vhdprovider.dll.mui",
	"sources/{mui_lang}/vofflps.rtf",
	"sources/{mui_lang}/vofflps_server.rtf",
	"sources/{mui_lang}/w32uires.dll.mui",
	"sources/{mui_lang}/wdsclient.dll.mui",
	"sources/{mui_lang}/wdsimage.dll.mui",
	"sources/{mui_lang}/wimgapi.dll.mui",
	"sources/{mui_lang}/wimprovider.dll.mui",
	"sources/{mui_lang}/WinDlp.dll.mui",
	"sources/{mui_lang}/winsetup.dll.mui"
]
EDITION_NAMES = {
    "core": "Home",
    "coren": "Home N",
    "professional": "Pro",
    "professionaln": "Pro N",
    "ppipro": "Team",
    "serverstandard": "Server Standard",
    "serverstandardcore": "Server Standard Core",
    "serverdatacenter": "Server Datacenter",
    "serverdatacentercore": "Server Datacenter Core"
}
EDITION_FLAGS = {
    "core": "Core",
    "coren": "CoreN",
    "professional": "Professional",
    "professionaln": "ProfessionalN",
    "ppipro": "PPIPro",
    "serverstandard": "ServerStandard",
    "serverstandardcore": "ServerStandardCor",
    "serverdatacenter": "ServerDatacenter",
    "serverdatacentercore": "ServerDatacenterCor"
}
BASE_EDITIONS = ["core", "coren", "serverstandard", "serverstandardcore", "ppipro"]
EDITION_BASES = {
    "professional": "core",
    "professionaln": "coren",
    "serverdatacenter": "serverstandard",
    "serverdatacentercore": "serverstandardcore"
}
EFI_BOOTS = {
    "x86": "bootia32.efi",
    "amd64": "bootx64.efi",
    "arm64": "bootaa64.efi"
}

def fetch_update_data(w, build, **kwargs):
    return list(filter(lambda u: re.search("Feature update|Upgrade to Windows \d+|Insider Preview|Windows \d+, version", u["title"], re.IGNORECASE), w.fetch_update_data(build, **kwargs)))

def extract(arc, fn, dirc):
    run(["7z", "e", "-y", arc, f"-o{dirc}", fn], stdout=PIPE)

def text_scan(text, pattern, flags=0):
    return re.findall(pattern, text, re.MULTILINE | re.IGNORECASE)

def wget(url, out, dirc, chksum):
    dirc = dirc.replace("\\", "/")
    run(["aria2c",  "--console-log-level=error", "--summary-interval=0", "--download-result=hide", "-c", f"--out={out}", f"--dir={dirc}", f"--checksum=sha-1={chksum}", url])

def wget_list(table):
    with open("dl.tmp", "w") as f:
        for row in table:
            if len(row) > 4:
                dirc = row[4]
            else:
                dirc = "UUP"
            
            f.write(f"{row[2]}\n\tout={row[0]}\n\tdir={dirc}\n\tchecksum=sha-1={row[1]}\n")
    
    run(["aria2c", "--console-log-level=error", "--summary-interval=0", "--download-result=hide", "-c", "-x", "16", "-j", "16", "-i", "dl.tmp"])
    remove("dl.tmp")

def search_updates(updates, file):
    fentry = list(next(filter(lambda f: file[0] in f[0] and f[3] == file[2], updates))) + list(file)[3:]
    fentry[0] = file[0]
    return fentry

def filter_updates(updates, name):
    return list(filter(lambda f: name in f[0] and "psf" not in f[0] and "baseless" not in f[0] and "express" not in f[0].lower() and ".msu" not in f[0], updates))

def wimlib_cmd(wim, index, cmd):
    run(["wimlib-imagex", "update", wim, index, "--command", cmd], stdout=PIPE)

def wimlib_cmds(wim, index, cmds):
    p = Popen(["wimlib-imagex", "update", wim, index], stdin=PIPE, stdout=PIPE)
    p.communicate("\n".join(cmds).encode("utf-8"))

def xcopy(src, dst, isdir=False):
    flags = ["/CIDRY", "/CEDRY"][isdir]
    run(["xcopy", flags, r"C:\mnt" + src, ISO_DIR + dst + "\\"], stdout=PIPE, stderr=PIPE)

if __name__ == "__main__":
    pcwd = getcwd()
    environ["PATH"] += rf";{pcwd}\bin"
    
    parser = argparse.ArgumentParser()
    parser.add_argument("version", nargs="?", default="insider", help="Windows version to download (default: insider)")
    parser.add_argument("--arch", "-a", default="amd64", help="Architecture of Windows (default: amd64)")
    parser.add_argument("--branch", "-b", default="vb_release", help="Manually specified branch (default: vb_release)")
    parser.add_argument("--ring", "-r", default="Retail", help="Manually specified ring (default: Retail)")
    parser.add_argument("--lang", "-l", default="en-us", help="Language of Windows (default: en-us)")
    parser.add_argument("--sku", "-s", default="Professional", help="SKU of product to download (default: Professional)")
    parser.add_argument("--editions", "-e", default=",".join(list(EDITION_NAMES)), help="Comma separated (no space) list of editions to create (default: all)")
    parser.add_argument("--pause-iso", "-p", help="Pause before ISO generation, useful for modded ISOs", action="store_true", default=False)
    parser.add_argument("--keep", "-k", help="Keep downloaded and temporary files (usually only needed for debugging)", action="store_true", default=False)
    parser.add_argument("--query", "-q", help="Only query updates, do not generate media", action="store_true", default=False)
    parser.add_argument("--dump", "-d", help="Only dump update information to JSON, do not generate media", action="store_true", default=False)
    args = parser.parse_args()
    
    editions = args.editions.lower().split(",")
    editions_dl = []
    lang = args.lang
    arch = args.arch
    version = args.version
    
    for edition in editions:
        if edition not in BASE_EDITIONS and EDITION_BASES[edition] not in editions:
            editions_dl = [EDITION_BASES[edition]] + editions_dl
        elif edition in BASE_EDITIONS:
            editions_dl.append(edition)
    
    print("Grabbing latest update info from Windows Update...")
    
    w = WUApi()
    
    if args.query:
        if version == "insider":
            version = "0"
        
        q_update_data = fetch_update_data(w, f"10.0.{version}.1", branch=args.branch, ring=args.ring, sku=args.sku)
        
        if len(q_update_data) == 0:
            print("Query returned no results.")
        
        for i, q_upd_item in enumerate(q_update_data):
            print(f"Update {i+1}:")
            print(f"Title: {q_upd_item['title']}")
            print(f"Full version: {q_upd_item['build']}")
            print(f"Update ID: {q_upd_item['id']}")
            print()
        
        exit(0)
    
    if version.lower() == "insider":
        uid = fetch_update_data(w, "10.0.0")[0]["id"]
    elif version.lower() == "10":
        uid = fetch_update_data(w, f"10.0.{LATEST_RETAIL}.1", branch="vb_release", ring="Retail")[0]["id"]
    elif version == "11":
        uid = fetch_update_data(w, f"10.0.22621.1", branch="ni_release", ring="Retail")[0]["id"]
    elif version.lower() in VERSION_BUILD:
        build, branch = VERSION_BUILD[version.lower()]
        uid = fetch_update_data(w, f"10.0.{build}.1", branch=branch, ring="Retail")[0]["id"]
    else:
        uid = fetch_update_data(w, f"10.0.{version}.1", branch=args.branch, ring=args.ring, sku=args.sku)[0]["id"]
    
    if args.dump:
        print(f"Dumping update info to {uid}.json")
        
        cache_entry = w.cache[uid]
        upd_files_json = [{
            "name": cache_entry["files"].get(hash, "N/A"),
            "sha1_hash": hash,
            "sha256_hash": cache_entry["ext_hashes"].get(hash, "N/A")
        } for hash in cache_entry["files"]]
        upd_files_json = sorted(upd_files_json, key=lambda f: f["name"])
        
        cache_entry["files"] = upd_files_json
        del cache_entry["ext_hashes"]
        
        with open(f"{uid}.json", "w") as jf:
            jf.write(dumps(cache_entry, indent=4))
        
        exit(0)
    
    upd_files = w.get_files(uid)
    aggr_meta = next(filter(lambda f: "AggregatedMetadata" in f[0], upd_files))
    aggr_fn = aggr_meta[0]
    
    build, spbuild = tuple(map(int, w.cache[uid]["build"].split(".")[2:]))
    full_title = w.cache[uid]["title"]
    
    try:
        win_title = re.search("Windows \d+", full_title)[0]
    except:
        win_title = "Windows"
    
    mui_lang = lang[:2] + lang[2:].upper()
    
    print(f"Processing update information for {full_title}...")
    
    dl_files = []
    iupd_table = []
    meta_esds = []
    appx_apps = []
    appx_licenses = {}
    appx_editions = defaultdict(list)
    
    with TemporaryDirectory() as tdir:
        chdir(tdir)
        wget(aggr_meta[2], aggr_fn, tdir, aggr_meta[1])
        match = "|".join(editions_dl)
        aggr_listing = run(["7z", "l", aggr_fn], stdout=PIPE).stdout.decode("utf-8")
        
        if build >= 22557 and not args.dump:
            for cabf in text_scan(aggr_listing, rf"\S+targetcompdb\S+(?:{match})_{lang}\.xml\.cab"):
                extract(aggr_fn, cabf, tdir)
                compdb_listing = run(["7z", "l", cabf], stdout=PIPE).stdout.decode("utf-8")
                xmlf = text_scan(compdb_listing, r"(\S+.xml)[^\.]")[0]
                extract(cabf, xmlf, tdir)
                compdb = parse_xml(open(xmlf).read())
                edition = re.match(rf"\S+targetcompdb\S+({match})_{lang}\.xml\.cab", cabf, re.IGNORECASE)[1]
                appxs = compdb.dependencies.find_all("feature")
                
                for appx in appxs:
                    appx_name = appx.attrs["featureid"]
                    
                    if edition in BASE_EDITIONS and appx_name not in appx_editions[edition]:
                        appx_editions[edition].append(appx_name)
                    
                    if appx_name not in appx_apps:
                        appx_apps.append(appx_name)
            
            for cabf in text_scan(aggr_listing, r"\S+targetcompdb_app\S+\.xml\.cab"):
                extract(aggr_fn, cabf, tdir)
                appdb_listing = run(["7z", "l", cabf], stdout=PIPE).stdout.decode("utf-8")
                xmlf = text_scan(appdb_listing, r"(\S+.xml)[^\.]")[0]
                extract(cabf, xmlf, tdir)
                appdb = parse_xml_html(open(xmlf).read())
                
                packages = appdb.compdb.find("packages", recursive=False)
                features = appdb.features
                
                frameworks = []
                
                for fmwk in features.find_all("feature", {"type": "MSIXFramework"}):
                    frameworks.append(fmwk.attrs["featureid"])
                    
                    for pkg in fmwk.packages.find_all("package"):
                        pkg_name = pkg.attrs["id"]
                        payload = packages.find("package", {"id": pkg_name}).payloaditem
                        
                        path = payload.attrs["path"].split("\\")
                        chksum = payload.attrs["payloadhash"]
                        fname = path[-1]
                        size = payload.attrs["payloadsize"]
                        fdata = (fname, size, b64hdec(chksum), "UUP/MSIXFramework")
                        
                        if fdata not in dl_files:
                            dl_files.append(fdata)
                
                for edition in appx_editions:
                    appx_editions[edition] = list(set(frameworks + appx_editions[edition]))
                
                for appx_name in appx_apps:
                    appx = features.find("feature", {"featureid": appx_name}, recursive=False)
                    
                    if appx is None:
                        continue
                    
                    print(f"Processing info for package {appx_name}...")
                    
                    if appx.custominfo:
                        appx_licenses[appx_name] = appx.custominfo.text
                    
                    for pkg in appx.packages.find_all("package"):
                        pkg_name = pkg.attrs["id"]
                        payload = packages.find("package", {"id": pkg_name}).payloaditem
                        
                        path = payload.attrs["path"].split("\\")
                        chksum = payload.attrs["payloadhash"]                    
                        fname = path[-1]
                        size = payload.attrs["payloadsize"]
                        
                        if appx.attrs["type"] == "MSIXFramework":
                            continue
                        elif "stub" in pkg.attrs["packagetype"].lower():
                            dirc = f"UUP/{appx_name}/AppxMetadata/Stub"
                        else:
                            dirc = f"UUP/{appx_name}"
                        
                        fdata = (fname, size, b64hdec(chksum), dirc)
                        
                        if fdata not in dl_files:
                            dl_files.append(fdata)
        
            chdir(pcwd)
            upd_files = w.get_files(uid)
            dl_table = list(map(lambda f: search_updates(upd_files, f), dl_files))
            
            if not exists(join(UUP_DIR, ".complete")):
                print()
                print("Downloading APPX files...")
                wget_list(dl_table)
                
                for appx in appx_licenses:
                    with open(join(UUP_DIR, appx, "License.xml"), "w") as f:
                        f.write(appx_licenses[appx])
            
            chdir(tdir)
            dl_files = []
        
        for cabf in text_scan(aggr_listing, rf"\S+targetcompdb\S*_(?:{match})_{lang}\.xml\.cab"):
            extract(aggr_fn, cabf, tdir)
            compdb_listing = run(["7z", "l", cabf], stdout=PIPE).stdout.decode("utf-8")
            xmlf = text_scan(compdb_listing, r"(\S+.xml)[^\.]")[0]
            extract(cabf, xmlf, tdir)
            compdb = parse_xml(open(xmlf).read())
            edition = re.match(rf"\S+targetcompdb\S+({match})_{lang}\.xml\.cab", cabf, re.IGNORECASE)[1]
            payloads = compdb.find_all("payloaditem")
            
            for payload in payloads:
                path = payload.attrs["path"].split("\\")
                chksum = payload.attrs["payloadhash"]
                size = payload.attrs["payloadsize"]
                
                if path[0] in ["UUP", "FeaturesOnDemand", "MetadataESD"]:
                    fname = path[-1]
                    fdata = (fname, size, b64hdec(chksum))
                    
                    if fdata not in dl_files:
                        dl_files.append(fdata)
                    
                    if path[0] == "MetadataESD":
                        meta_esds.append(fname)
        
        chdir(pcwd)
        upd_files = w.get_files(uid)
        kb_upd = filter_updates(upd_files, "Windows11.0-KB") + filter_updates(upd_files, "Windows10.0-KB")
        ssu_upd = filter_updates(upd_files, "SSU")
        iupd_table += sorted(ssu_upd + kb_upd, key=lambda f: f[0])
        dl_table = []
        dl_table = list(map(lambda f: search_updates(upd_files, f), dl_files))
        
        if not exists(join(UUP_DIR, ".complete")):
            print()
            print("Downloading files...")
            wget_list(dl_table + iupd_table)
            open(join(UUP_DIR, ".complete"), "w").close()
    
    ref_esds_exist = False
    if not exists(TEMP_DIR):
        print()
        print("Preparing reference ESDs...")
        
        makedirs(TEMP_DIR, exist_ok=True)
        for fentry in dl_table:
            pkg = fentry[0]
            
            if pkg.endswith(".cab"):
                ref_esds_exist = True
                print(f"Converting {pkg}...")
                
                with TemporaryDirectory(dir=TEMP_DIR) as ctdir:
                    run(["expand.exe", "-f:*", join(UUP_DIR, pkg), f"{ctdir}\\"], stdout=PIPE)
                    run(["wimlib-imagex", "capture", ctdir, join(TEMP_DIR, pkg + ".ESD"), "--compress=XPRESS", "--check", "--no-acls", "--norpfix", "Edition Package", "Edition Package"], stdout=PIPE)
                    run(["cmd", "/c", "rmdir", "/s", "/q", ctdir])
                    makedirs(ctdir)
    
    print("Creating ISODIR...")
    
    meta_base = join(UUP_DIR, meta_esds[0])
    makedirs(ISO_DIR, exist_ok=True)
    run(["wimlib-imagex", "apply", meta_base, "1", ISO_DIR, "--no-acls", "--no-attributes"])
    
    print("Exporting WinRE...")
    
    winre_wim = join(TEMP_DIR, "winre.wim")
    run(["wimlib-imagex", "export", meta_base, "2", winre_wim, "--compress=LZX", "--boot"])
    
    print("Configuring Boot WIM...")
    
    boot_wim = join(ISO_DIR, "sources", "boot.wim")
    copyfile(winre_wim, boot_wim)
    run(["wimlib-imagex", "info", boot_wim, "1", "Microsoft Windows PE", "Microsoft Windows PE", "--image-property", "FLAGS=9"], stdout=PIPE)
    run(["wimlib-imagex", "extract", boot_wim, "1", f"--dest-dir={TEMP_DIR}", "/Windows/System32/config/SOFTWARE", "--no-acls", "--no-attributes"], stdout=PIPE)
    run(["offlinereg", join(TEMP_DIR, "SOFTWARE"), r"Microsoft\Windows NT\CurrentVersion\WinPE", "setvalue", "InstRoot", "X:\\$windows.~bt\\"], stdout=PIPE)
    run(["offlinereg", join(TEMP_DIR, "SOFTWARE.new"), r"Microsoft\Windows NT\CurrentVersion", "setvalue", "SystemRoot", r"X:\$windows.~bt\Windows"], stdout=PIPE)
    remove(join(TEMP_DIR, "SOFTWARE"))
    move(join(TEMP_DIR, "SOFTWARE.new"), join(TEMP_DIR, "SOFTWARE"))
    wimlib_cmd(boot_wim, "1", f"add {join(TEMP_DIR, 'SOFTWARE')} /Windows/System32/config/SOFTWARE")
    remove(join(TEMP_DIR, 'SOFTWARE'))
    
    for bgfn in ["background_cli.bmp", "background_cli.png", "background_svr.bmp", "background_svr.png"]:
        if exists(join(ISO_DIR, 'sources', bgfn)):
            bgcli = join(ISO_DIR, 'sources', bgfn)
            bgimg = "/sources/background." + bgfn.split(".")[-1]
    
    wimlib_cmd(boot_wim, "1", f"add {bgcli} /Windows/System32/winre.jpg")
    wimlib_cmd(boot_wim, "1", f"delete /Windows/System32/winpeshl.ini")
    run(["wimlib-imagex", "export", winre_wim, "1", boot_wim, "Microsoft Windows Setup", "Microsoft Windows Setup"], stdout=PIPE)
    run(["wimlib-imagex", "extract", meta_base, "3", "/Windows/System32/xmllite.dll", f"--dest-dir={join(ISO_DIR, 'sources')}", "--no-acls", "--no-attributes"], stdout=PIPE)
    run(["wimlib-imagex", "info", boot_wim, "2", "--image-property", "FLAGS=2"], stdout=PIPE)
    run(["wimlib-imagex", "info", boot_wim, "2", "--boot"], stdout=PIPE)
    
    src_files = list(map(lambda f: f.format(mui_lang=mui_lang), filter(lambda f: exists(join(ISO_DIR, f.format(mui_lang=mui_lang))), BOOT_SRC)))
    src_cmds = [
        "delete /Windows/System32/winpeshl.ini",
        f"add {join(ISO_DIR, 'setup.exe')} /setup.exe",
        f"add {join(ISO_DIR, 'sources', 'inf', 'setup.cfg')} /sources/inf/setup.cfg",
        f"add {bgcli} {bgimg}",
        f"add {bgcli} /Windows/system32/winre.jpg"
    ]
    src_cmds += [f"add {join(ISO_DIR, f)} /{f}" for f in src_files]
    wimlib_cmds(boot_wim, "2", src_cmds)
    
    run(["wimlib-imagex", "optimize", boot_wim], stdout=PIPE)
    remove(join(ISO_DIR, 'sources', 'xmllite.dll'))
    
    print("Preparing install WIM...")
    
    install_wim = join(ISO_DIR, "sources", "install.wim")
    makedirs(r"C:\mnt", exist_ok=True)
    makedirs(r"C:\uup", exist_ok=True)
    
    index = 1
    inst_editions = {}
    copied_efi = False
    
    for edition in editions:
        ed_name = EDITION_NAMES[edition]
        
        if edition in BASE_EDITIONS:
            try:
                meta_esd = join(UUP_DIR, next(filter(lambda m: m.lower().startswith(f"{edition}_"), meta_esds)))
            except:
                continue
            
            print(f"Creating edition {win_title} {ed_name}")
            run(["wimlib-imagex", "export", meta_esd, "3", install_wim, f"{win_title} {ed_name}", r"--ref=UUP\*.esd", r"--ref=TEMP\*.esd" if ref_esds_exist else "", "--compress=LZX"])
            run(["dism", r"/scratchdir:C:\uup", "/mount-wim", "/wimfile:" + install_wim.replace("/", "\\"), f"/index:{index}", r"/mountdir:C:\mnt"])
            
            for upd in iupd_table:
                upd_fname = abspath(join(UUP_DIR, upd[0]).replace("/", "\\"))
                run(["dism", r"/scratchdir:C:\uup", r"/image:C:\mnt", f"/add-package:{upd_fname}"])
                
                if "SSU" in upd[0]:
                    run(["dism", r"/scratchdir:C:\uup", "/unmount-image", r"/mountdir:C:\mnt", "/commit"])
                    run(["dism", r"/scratchdir:C:\uup", "/mount-wim", "/wimfile:" + install_wim.replace("/", "\\"), f"/index:{index}", r"/mountdir:C:\mnt"])
            
            if build >= 22557 and exists(join(UUP_DIR, "MSIXFramework")):
                print("Installing base app libraries...")
                
                for appx_file in listdir(join(UUP_DIR, "MSIXFramework")):
                    appx_path = abspath(join(UUP_DIR, "MSIXFramework", appx_file).replace("/", "\\"))
                    run(["dism", r"/scratchdir:C:\uup", r"/image:C:\mnt", "/add-provisionedappxpackage", f"/packagepath:{appx_path}", "/skiplicense"])
                
                print("Installing apps...")
                
                for appx_folder in filter(lambda f: isdir(join(UUP_DIR, f)), appx_editions[edition]):
                    try:
                        pkg_file = next(filter(lambda f: "bundle" in f.lower(), listdir(join(UUP_DIR, appx_folder))))
                    except:
                        pkg_file = next(filter(lambda f: "appx" in f.lower() or "msix" in f.lower(), listdir(join(UUP_DIR, appx_folder))))
                    
                    appx_path = abspath(join(UUP_DIR, appx_folder, pkg_file).replace("/", "\\"))
                    lic_path = abspath(join(UUP_DIR, appx_folder, "License.xml").replace("/", "\\"))
                    stub = []
                    
                    if exists(join(UUP_DIR, appx_folder, "AppxMetadata", "Stub")):
                        stub = ["/stubpackageoption:installstub"]
                    
                    print(f"Installing app {appx_folder}...")
                    run(["dism", r"/scratchdir:C:\uup", r"/image:C:\mnt", "/add-provisionedappxpackage", f"/packagepath:{appx_path}", f"/licensepath:{lic_path}"] + stub)
            
            if not copied_efi:
                efi_file = EFI_BOOTS[arch]
                
                if arch != "arm64":
                    xcopy(r"\Windows\Boot\PCAT\bootmgr", "")
                    xcopy(r"\Windows\Boot\PCAT\memtest.exe", r"\boot")
                    xcopy(r"\Windows\Boot\EFI\memtest.efi", r"\efi\microsoft\boot")
                
                xcopy(r"\Windows\Boot\EFI\bootmgfw.efi", rf"\efi\boot\{efi_file}")
                xcopy(r"\Windows\Boot\EFI\bootmgr.efi", "")
                
                if exists(r"C:\mnt\Windows\Boot\EFI\winsipolicy.p7b"):
                    xcopy(r"\Windows\Boot\EFI\winsipolicy.p7b", r"\efi\microsoft\boot\winsipolicy.p7b", True)
                if exists(r"C:\mnt\Windows\Boot\EFI\CIPolicies"):
                    xcopy(r"\Windows\Boot\EFI\CIPolicies", r"\efi\microsoft\boot\cipolicies", True)
                
                if build >= 18890:
                    xcopy(r"\Windows\Boot\Fonts", r"\boot\fonts")
                    xcopy(r"\Windows\Boot\Fonts", r"\efi\microsoft\boot\fonts")
                
                copied_efi = True
            
            run(["dism", r"/scratchdir:C:\uup", "/unmount-image", r"/mountdir:C:\mnt", "/commit"])
        else:
            base_edition = EDITION_BASES[edition]
            
            if base_edition not in inst_editions:
                continue
            
            print(f"Creating edition {win_title} {ed_name}")
            run(["dism", r"/scratchdir:C:\uup", "/mount-wim", "/wimfile:" + install_wim.replace("/", "\\"), f"/index:{inst_editions[base_edition]}", r"/mountdir:C:\mnt"])
            run(["dism", r"/scratchdir:C:\uup", "/image:C:\mnt", f"/set-edition:{EDITION_FLAGS[edition]}", "/channel:retail"])
            run(["dism", r"/scratchdir:C:\uup", "/unmount-image", r"/mountdir:C:\mnt", "/commit", "/append"])
        
        desc = f"{win_title} {ed_name}"
        run(["wimlib-imagex", "info", install_wim, str(index), desc, desc, "--image-property", f"DISPLAYNAME={desc}", "--image-property", f"DISPLAYDESCRIPTION={desc}", "--image-property", f"FLAGS={EDITION_FLAGS[edition]}"])
        wimlib_cmd(install_wim, str(index), f"add {winre_wim} /Windows/System32/recovery/winre.wim")
        inst_editions[edition] = index
        index += 1
    
    run(["wimlib-imagex", "optimize", install_wim], stdout=PIPE)
    remove(winre_wim)
    
    if args.pause_iso:
        input("Press enter to generate ISO")
    
    print("Generating ISO...")
    
    timestamp = datetime.now().strftime("%y%m%d-%H%M")
    branch = w.cache[uid]["branch"]
    label = f"WIN{build}_{arch.upper()}FRE_{lang.upper()}"
    
    edition_label = "MULTI" if len(editions) > 1 else editions[0]
    cs_label = "SERVER" if any(["server" in meta_esd for meta_esd in meta_esds]) else "CLIENT"
    
    filename = f"{build}.{spbuild}.{timestamp}.{branch}_{arch.upper()}_{edition_label}_{cs_label}_{lang}.ISO"
    
    if arch != "arm64":
        run(["cdimage", rf'-bootdata:2#p0,e,b{ISO_DIR}\boot\etfsboot.com#pEF,e,b{ISO_DIR}\efi\Microsoft\boot\efisys.bin', "-o", "-m", "-u2", "-udfver102", f"-l{label}", ISO_DIR, filename])
    else:
        run(["cdimage", rf'-bootdata:1#pEF,e,b{ISO_DIR}\efi\Microsoft\boot\efisys.bin', "-o", "-m", "-u2", "-udfver102", f"-l{label}", ISO_DIR, filename])
    
    print("Cleaning up...")
    
    if not args.keep:
        rmtree(UUP_DIR)
    
    rmtree(TEMP_DIR)
    rmtree(r"C:\mnt")
    rmtree(r"C:\uup")
    rmtree(ISO_DIR)