import msal
import atexit
import requests
import time
import datetime
import os
import ping3
import magichue


# light = False
light = magichue.Light("<BULB IP ADDRESS>")

start = datetime.time(6, 0, 0)
end = datetime.time(17, 30, 0)

token = ""
CLIENT_ID = ""
TENANT_ID = ""
SECRET_ID = ""

AUTHORITY = "https://login.microsoftonline.com/" + TENANT_ID
ENDPOINT = "https://graph.microsoft.com/v1.0"
SCOPES = ["Presence.Read"]

# config = configparser.ConfigParser()
# with open("azure_config.ini", "w") as configfile:
#     config.write(configfile)


def Authorize():
    global token
    global fullname
    print("Starting authentication workflow.")
    try:
        cache = msal.SerializableTokenCache()
        if os.path.exists("token_cache.bin"):
            cache.deserialize(open("token_cache.bin", "r").read())

        atexit.register(
            lambda: open("token_cache.bin", "w").write(cache.serialize())
            if cache.has_state_changed
            else None
        )

        app = msal.PublicClientApplication(
            CLIENT_ID, authority=AUTHORITY, token_cache=cache
        )

        accounts = app.get_accounts()
        result = None
        if len(accounts) > 0:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])

        if result is None:
            # Create QR code
            # qr = pyqrcode.create("https://microsoft.com/devicelogin")
            # print(qr.terminal(module_color=0, background=231, quiet_zone=1))

            # Initiate flow
            flow = app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                raise Exception("Failed to create device flow")
            print(flow["message"])
            result = app.acquire_token_by_device_flow(flow)
            print(f"RESULT: {result}")
            token = result["access_token"]

            print("Aquired token")
            token_claim = result["id_token_claims"]
            print("Welcome " + token_claim.get("name") + "!")
            fullname = token_claim.get("name")
            return True
        if "access_token" in result:
            token = result["access_token"]
            try:
                result = requests.get(
                    f"{ENDPOINT}/me",
                    headers={"Authorization": "Bearer " + result["access_token"]},
                    timeout=5,
                )
                result.raise_for_status()
                y = result.json()
                fullname = y["givenName"] + " " + y["surname"]
                print("Token found, welcome " + y["givenName"] + "!")
                return True
            except requests.exceptions.HTTPError as err:
                if err.response.status_code == 404:
                    print("MS Graph URL is invalid!")
                    exit(5)
                elif err.response.status_code == 401:
                    print(
                        "MS Graph is not authorized. Please reauthorize the app (401)."
                    )
                    return False
            except requests.exceptions.Timeout as timeerr:
                print("The authentication request timed out. " + str(timeerr))
        else:
            raise Exception("no access token in result")
    except Exception as e:
        print("Failed to authenticate. " + str(e))
        time.sleep(2)
        return False


def time_in_range(start, end, current):
    """Returns whether current is in the range [start, end]"""
    return start <= current <= end


if __name__ == "__main__":
    # Tell Python to run the handler() function when SIGINT is recieved
    # signal(SIGINT, handler)

    trycount = 0
    while Authorize() == False:
        trycount = trycount + 1
        if trycount > 10:
            print("Cannot authorize. Will exit.")
            exit(5)
        else:
            print(
                "Failed authorizing, empty token ("
                + str(trycount)
                + "/10). Will try again in 10 seconds."
            )
            Authorize()
            continue

    time.sleep(1)

    trycount = 0

    while True:
        current = datetime.datetime.now().time()

        # try:
        #     light = magichue.Light("192.168.0.83")
        # except Exception as e:
        #     print("\033[91m" + "Bulb Offline" + "\033[0m")

        if time_in_range(start, end, current) == False:
            print(f"Bulb Offline at {datetime.datetime.now().strftime('%H:%M:%S')}")
            light.on = False
            time.sleep(60)
            continue

        print(
            f"\n ***Fetching new data at {datetime.datetime.now().strftime('%m-%d-%y %H:%M:%S')}***"
        )

        headers = {"Authorization": "Bearer " + token}
        jsonresult = ""

        try:
            result = requests.get(
                f"https://graph.microsoft.com/v1.0/me/presence",
                headers=headers,
                timeout=5,
            )
            result.raise_for_status()
            jsonresult = result.json()
            print(jsonresult)

        except requests.exceptions.Timeout as timeerr:
            print("The request for Graph API timed out! " + str(timeerr))
            continue

        except requests.exceptions.HTTPError as err:
            if err.response.status_code == 404:
                print("MS Graph URL is invalid!")
                exit(5)
            elif err.response.status_code == 401:
                trycount = trycount + 1
                print(
                    "MS Graph is not authorized. Please reauthorize the app (401). Trial count: "
                    + str(trycount)
                )
                print()
                Authorize()
                continue
        except:
            print("Will try again. Trial count: " + str(trycount))
            print()
            continue

        trycount = 0

        # Check for jsonresult
        if jsonresult == "":
            print("JSON result is empty! Will try again.")
            print(jsonresult)
            time.sleep(5)
            continue

        if jsonresult["activity"] == "Available":
            print("Teams presence:\t\t" + "\033[32m" + "Available" + "\033[0m")

            # if not light.is_white:
            #     light.is_white = True
            # light.cw = 0
            # light.w = 255
            # if light.on:
            light.on = False

            print("Light On: \t\t" + "\033[91m" + f"{light.on}" + "\033[0m")

        elif (
            jsonresult["activity"] == "Busy"
            or jsonresult["activity"] == "DoNotDisturb"
            or jsonresult["activity"] == "InAMeeting"
        ):
            print("Teams presence:\t\t" + "\033[91m" + "In Meeting" + "\033[0m")

            if not light.on:
                light.on = True

            if light.is_white:
                light.is_white = False

            light.rgb = (255, 0, 0)

            print("Light On: \t\t" + "\033[32m" + f"{light.on}" + "\033[0m")

        print()
        time.sleep(30)
