from random import randbytes
from binascii import unhexlify
from utils import b64benc, b64senc, b64hdec, wu_request, header_data, parse_xml
from consts import FLIGHT_BRANCHES, SKU_IDS, DVC_FAMILIES, INS_TYPES, PRODUCTS
from time import time
from html import escape
from bs4 import BeautifulSoup

class WUApi:
    _device = None
    _cookie = None
    cache = {}
    
    @property
    def device(self):
        if self._device is None:
            tbytes = b"\x13\x000\x02\xc3w\x04\x00\x14\xd5\xbc\xaczf\xde\rP\xbe\xdd\xf9\xbb\xa1l\x87\xed\xb9\xe0\x19\x89\x80\x00"
            tbytes += randbytes(527)
            tbytes += b"\xb4\x01"
            
            t = b64benc(tbytes)
            dev_str = f"t={t}&p="
            
            self._device = b64senc("".join([x + "\0" for x in dev_str]))
        
        return self._device
    
    @property
    def cookie(self):
        if self._cookie is None:
            uuid, create_date, expire_date = header_data()
            post_data = f"""
<s:Envelope xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:s="http://www.w3.org/2003/05/soap-envelope">
    <s:Header>
        <a:Action s:mustUnderstand="1">http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService/GetCookie</a:Action>
        <a:MessageID>urn:uuid:{uuid}</a:MessageID>
        <a:To s:mustUnderstand="1">https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx</a:To>
        <o:Security s:mustUnderstand="1" xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
            <Timestamp xmlns="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
                <Created>{create_date}</Created>
                <Expires>{expire_date}</Expires>
            </Timestamp>
            <wuws:WindowsUpdateTicketsToken wsu:id="ClientMSA" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd" xmlns:wuws="http://schemas.microsoft.com/msus/2014/10/WindowsUpdateAuthorization">
                <TicketType Name="MSA" Version="1.0" Policy="MBI_SSL">
                    <Device>{self.device}</Device>
                </TicketType>
            </wuws:WindowsUpdateTicketsToken>
        </o:Security>
    </s:Header>
    <s:Body>
        <GetCookie xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService">
            <oldCookie>
                <Expiration>{create_date}</Expiration>
            </oldCookie>
            <lastChange>{create_date}</lastChange>
            <currentTime>{create_date}</currentTime>
            <protocolVersion>2.0</protocolVersion>
        </GetCookie>
    </s:Body>
</s:Envelope>
"""
            
            resp = wu_request("https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx", post_data)
            #print(resp)
            #set_trace()
            self._cookie = {}
            
            if not (resp.find("NewCookie") or resp.find("GetCookieResult")) is None:
                self._cookie["expiration"] = resp.Expiration.text
                self._cookie["data"] = resp.EncryptedData.text
            else:
                raise Exception("Could not retrieve cookie")
        
        return self._cookie
    
    def device_attributes(self, build, branch, ring, arch, sku, flight, rs_type, **kwargs):
        ring = ring.upper()
        sku_id = SKU_IDS[sku]
        block_upgrades = int(sku_id in [125, 126, 7, 8, 12, 13, 79, 80, 120, 145, 146, 147, 148, 159, 160, 406, 407, 408])
        dvc_family = DVC_FAMILIES.get(sku_id, "Windows.Desktop")
        ins_type = INS_TYPES.get(sku_id, "Client")
        flight_enabled = 1
        is_retail = 0
        flight_branch = FLIGHT_BRANCHES[ring]
        flight_ring = "External"
        
        if ring == "RETAIL":
            flight_ring = "Retail"
            flight_enabled = 0
            is_retail = 1
        elif ring == "RP" and flight == "Active":
            flight = "Current"
        
        return escape("E:" + "&".join([
                "App=WU_OS",
                "AppVer=" + build,
                "AttrDataVer=134",
                "BlockFeatureUpdates=" + str(block_upgrades),
                "BranchReadinessLevel=CB",
                "CurrentBranch=" + branch,
                "DataExpDateEpoch_21H1=" + str(int(time()) + 82800),
                "DataExpDateEpoch_20H1=" + str(int(time()) + 82800),
                "DataExpDateEpoch_19H1=" + str(int(time()) + 82800),
                "DataVer_RS5=2000000000",
                "DefaultUserRegion=191",
                "DeviceFamily=" + dvc_family,
                "EKB19H2InstallCount=1",
                "EKB19H2InstallTimeEpoch=1255000000",
                "FlightingBranchName=" + flight_branch,
                "FlightRing=" + flight_ring,
                "Free=32to64",
                "GStatus_CO21H2=2",
                "GStatus_21H1=2",
                "GStatus_20H1=2",
                "GStatus_20H1Setup=2",
                "GStatus_19H1=2",
                "GStatus_19H1Setup=2",
                "GStatus_RS5=2",
                "GenTelRunTimestamp_19H1=" + str(int(time()) - 3600),
                "InstallDate=1438196400",
                "InstallLanguage=en-US",
                "InstallationType=" + ins_type,
                "IsDeviceRetailDemo=0",
                "IsFlightingEnabled=" + str(flight_enabled),
                "IsRetailOS=" + str(is_retail),
                "MediaBranch=",
                "MediaVersion=" + build,
                "CloudPBR=1",
                "DUScan=1",
                "OEMModel=Nonexistent Computer 696969",
                "OEMModelBaseBoard=Nonexistent Board",
                "OEMName_Uncleaned=Nonexistent Technologies LLC",
                "OSArchitecture=" + arch,
                "OSSkuId=" + str(sku_id),
                "OSUILocale=en-US",
                "OSVersion=" + build,
                "ProcessorIdentifier=Intel64 Family 6 Model 85 Stepping 4",
                "ProcessorManufacturer=GenuineIntel",
                "ReleaseType=" + rs_type,
                "SdbVer_20H1=2000000000",
                "SdbVer_19H1=2000000000",
                "TelemetryLevel=3",
                "TimestampEpochString_21H1=" + str(int(time()) - 3600),
                "TimestampEpochString_20H1=" + str(int(time()) - 3600),
                "TimestampEpochString_19H1=" + str(int(time()) - 3600),
                "UpdateManagementGroup=2",
                "UpdateOfferedDays=0",
                "UpgEx_CO21H2=Green",
                "UpgEx_21H1=Green",
                "UpgEx_20H1=Green",
                "UpgEx_19H1=Green",
                "UpgEx_RS5=Green",
                "UpgradeEligible=1",
                "Version_RS5=2000000000",
                "WuClientVer=" + build
        ]))
    
    def fetch_update_data(self, build, branch="co_release", ring="WIF", arch="amd64", sku="Professional", flight="Active", rs_type="Production"):
        uuid, create_date, expire_date = header_data()
        sku_id = SKU_IDS[sku]
        main_product = PRODUCTS.get(sku_id, "Client.OS.rs2")
        
        products = escape(";".join([
            f"PN={main_product}.{arch}&Branch={branch}&PrimaryOSProduct=1&Repairable=1&V={build}&ReofferUpdate=1",
            f"PN=Windows.Appraiser.{arch}&Repairable=1&V={build}",
            f"PN=Windows.AppraiserData.{arch}&Repairable=1&V={build}",
            f"PN=Windows.EmergencyUpdate.{arch}&Repairable=1&V={build}",
            f"PN=Windows.FeatureExperiencePack.{arch}&Repairable=1&V=0.0.0.0",
            f"PN=Windows.ManagementOOBE.{arch}&IsWindowsManagementOOBE=1&Repairable=1&V={build}",
            f"PN=Windows.OOBE.{arch}&IsWindowsOOBE=1&Repairable=1&V={build}",
            f"PN=Windows.UpdateStackPackage.{arch}&Name=Update Stack Package&Repairable=1&V={build}",
            f"PN=Hammer.{arch}&Source=UpdateOrchestrator&V=0.0.0.0",
            f"PN=MSRT.{arch}&Source=UpdateOrchestrator&V=0.0.0.0",
            f"PN=SedimentPack.{arch}&Source=UpdateOrchestrator&V=0.0.0.0",
            f"PN=Microsoft.Edge.Stable.{arch}&Repairable=1&V=0.0.0.0",
            f"PN=Adobe.Flash.{arch}&Repairable=1&V={build}",
            f"PN=Microsoft.NETFX.{arch}&V=2018.12.2.0"
        ]))
        
        caller_attrs = escape("E:" + "&".join([
            "Profile=AUv2",
            "Acquisition=1",
            "Interactive=1",
            "IsSeeker=1",
            "SheddingAware=1",
            "Id=MoUpdateOrchestrator"
        ]))
        
        device_attrs = self.device_attributes(build, branch, ring, arch, sku, flight, rs_type)
        
        post_data = f"""
<s:Envelope xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:s="http://www.w3.org/2003/05/soap-envelope">
    <s:Header>
        <a:Action s:mustUnderstand="1">http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService/SyncUpdates</a:Action>
        <a:MessageID>urn:uuid:{uuid}</a:MessageID>
        <a:To s:mustUnderstand="1">https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx</a:To>
        <o:Security s:mustUnderstand="1" xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
            <Timestamp xmlns="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
                <Created>{create_date}</Created>
                <Expires>{expire_date}</Expires>
            </Timestamp>
            <wuws:WindowsUpdateTicketsToken wsu:id="ClientMSA" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd" xmlns:wuws="http://schemas.microsoft.com/msus/2014/10/WindowsUpdateAuthorization">
                <TicketType Name="MSA" Version="1.0" Policy="MBI_SSL">
                    <Device>{self.device}</Device>
                </TicketType>
            </wuws:WindowsUpdateTicketsToken>
        </o:Security>
    </s:Header>
    <s:Body>
        <SyncUpdates xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService">
            <cookie>
                <Expiration>{self.cookie["expiration"]}</Expiration>
                <EncryptedData>{self.cookie["data"]}</EncryptedData>
            </cookie>
            <parameters>
                <ExpressQuery>false</ExpressQuery>
                <InstalledNonLeafUpdateIDs>
                    <int>1</int>
                    <int>10</int>
                    <int>105939029</int>
                    <int>105995585</int>
                    <int>106017178</int>
                    <int>107825194</int>
                    <int>10809856</int>
                    <int>11</int>
                    <int>117765322</int>
                    <int>129905029</int>
                    <int>130040030</int>
                    <int>130040031</int>
                    <int>130040032</int>
                    <int>130040033</int>
                    <int>133399034</int>
                    <int>138372035</int>
                    <int>138372036</int>
                    <int>139536037</int>
                    <int>139536038</int>
                    <int>139536039</int>
                    <int>139536040</int>
                    <int>142045136</int>
                    <int>158941041</int>
                    <int>158941042</int>
                    <int>158941043</int>
                    <int>158941044</int>
                    <int>159776047</int>
                    <int>160733048</int>
                    <int>160733049</int>
                    <int>160733050</int>
                    <int>160733051</int>
                    <int>160733055</int>
                    <int>160733056</int>
                    <int>161870057</int>
                    <int>161870058</int>
                    <int>161870059</int>
                    <int>17</int>
                    <int>19</int>
                    <int>2</int>
                    <int>23110993</int>
                    <int>23110994</int>
                    <int>23110995</int>
                    <int>23110996</int>
                    <int>23110999</int>
                    <int>23111000</int>
                    <int>23111001</int>
                    <int>23111002</int>
                    <int>23111003</int>
                    <int>23111004</int>
                    <int>2359974</int>
                    <int>2359977</int>
                    <int>24513870</int>
                    <int>28880263</int>
                    <int>3</int>
                    <int>30077688</int>
                    <int>30486944</int>
                    <int>5143990</int>
                    <int>5169043</int>
                    <int>5169044</int>
                    <int>5169047</int>
                    <int>59830006</int>
                    <int>59830007</int>
                    <int>59830008</int>
                    <int>60484010</int>
                    <int>62450018</int>
                    <int>62450019</int>
                    <int>62450020</int>
                    <int>69801474</int>
                    <int>8788830</int>
                    <int>8806526</int>
                    <int>9125350</int>
                    <int>9154769</int>
                    <int>98959022</int>
                    <int>98959023</int>
                    <int>98959024</int>
                    <int>98959025</int>
                    <int>98959026</int>
                </InstalledNonLeafUpdateIDs>
                <OtherCachedUpdateIDs/>
                <SkipSoftwareSync>false</SkipSoftwareSync>
                <NeedTwoGroupOutOfScopeUpdates>true</NeedTwoGroupOutOfScopeUpdates>
                <AlsoPerformRegularSync>true</AlsoPerformRegularSync>
                <ComputerSpec/>
                <ExtendedUpdateInfoParameters>
                    <XmlUpdateFragmentTypes>
                        <XmlUpdateFragmentType>Extended</XmlUpdateFragmentType>
                        <XmlUpdateFragmentType>LocalizedProperties</XmlUpdateFragmentType>
                    </XmlUpdateFragmentTypes>
                    <Locales>
                        <string>en-US</string>
                    </Locales>
                </ExtendedUpdateInfoParameters>
                <ClientPreferredLanguages/>
                <ProductsParameters>
                    <SyncCurrentVersionOnly>false</SyncCurrentVersionOnly>
                    <DeviceAttributes>{device_attrs}</DeviceAttributes>
                    <CallerAttributes>{caller_attrs}</CallerAttributes>
                    <Products>{products}</Products>
                </ProductsParameters>
            </parameters>
        </SyncUpdates>
    </s:Body>
</s:Envelope>
"""
        
        resp = wu_request("https://fe3cr.delivery.mp.microsoft.com/ClientWebService/client.asmx", post_data)
        
        updates = resp.NewUpdates.find_all("UpdateInfo")
        results = []
        
        for u in updates:
            result = {}
            files = {}
            ext_hashes = {}
            
            update_meta = parse_xml(u.Xml.text)
            update_ext_meta = parse_xml(resp.find("ID", string=lambda s: s == u.ID.text and "LocalizedProperties" in s.parent.parent.Xml.text).parent.Xml.text)
            update_ext_props = parse_xml(resp.find("ID", string=lambda s: s == u.ID.text and "ExtendedProperties" in s.parent.parent.Xml.text).parent.Xml.text)
            
            update_id = update_meta.updateidentity.attrs["updateid"]
            update_title = update_ext_meta.title.text
            update_version = update_meta.productreleaseinstalled.attrs["version"]
            
            result["id"] = update_id
            result["title"] = update_title
            result["build"] = update_version
            
            for file in update_ext_props.find_all("file"):
                files[b64hdec(file.attrs["digest"])] = file.attrs["filename"]
                ext_hashes[file.attrs["filename"]] = b64hdec(file.additionaldigest.text)
            
            self.cache[update_id] = {
                "arch": arch,
                "build": update_version,
                "branch": branch,
                "ext_hashes": ext_hashes,
                "files": files,
                "flight": flight,
                "ring": ring,
                "rs_type": rs_type,
                "sku": sku,
                "title": update_title
            }
            
            results.append(result)
        
        return results
    
    def get_files(self, update_id, rev=1):
        uuid, create_date, expire_date = header_data()
        cache_data = self.cache[update_id]
        device_attrs = self.device_attributes(**cache_data)
        
        post_data = f"""
<s:Envelope xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:s="http://www.w3.org/2003/05/soap-envelope">
    <s:Header>
        <a:Action s:mustUnderstand="1">http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService/GetExtendedUpdateInfo2</a:Action>
        <a:MessageID>urn:uuid:{uuid}</a:MessageID>
        <a:To s:mustUnderstand="1">https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx/secured</a:To>
        <o:Security s:mustUnderstand="1" xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
            <Timestamp xmlns="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
                <Created>{create_date}</Created>
                <Expires>{expire_date}</Expires>
            </Timestamp>
            <wuws:WindowsUpdateTicketsToken wsu:id="ClientMSA" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd" xmlns:wuws="http://schemas.microsoft.com/msus/2014/10/WindowsUpdateAuthorization">
                <TicketType Name="MSA" Version="1.0" Policy="MBI_SSL">
                    <Device>{self.device}</Device>
                </TicketType>
            </wuws:WindowsUpdateTicketsToken>
        </o:Security>
    </s:Header>
    <s:Body>
        <GetExtendedUpdateInfo2 xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService">
            <updateIDs>
                <UpdateIdentity>
                    <UpdateID>{update_id}</UpdateID>
                    <RevisionNumber>{rev}</RevisionNumber>
                </UpdateIdentity>
            </updateIDs>
            <infoTypes>
                <XmlUpdateFragmentType>FileUrl</XmlUpdateFragmentType>
                <XmlUpdateFragmentType>FileDecryption</XmlUpdateFragmentType>
                <XmlUpdateFragmentType>EsrpDecryptionInformation</XmlUpdateFragmentType>
                <XmlUpdateFragmentType>PiecesHashUrl</XmlUpdateFragmentType>
                <XmlUpdateFragmentType>BlockMapUrl</XmlUpdateFragmentType>
            </infoTypes>
            <deviceAttributes>{device_attrs}</deviceAttributes>
        </GetExtendedUpdateInfo2>
    </s:Body>
</s:Envelope>
"""
        
        resp = wu_request("https://fe3cr.delivery.mp.microsoft.com/ClientWebService/client.asmx/secured", post_data)
        result = []
        
        for file in resp.find_all("FileLocation"):
            fhash = b64hdec(file.FileDigest.text)
            fname = self.cache[update_id]["files"].get(fhash, "")
            furl = file.Url.text
            fehash = self.cache[update_id]["ext_hashes"][fname]
            
            if fname != "":
                result.append((fname, fhash, furl, fehash))
        
        return result
