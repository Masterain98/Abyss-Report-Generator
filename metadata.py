import httpx


def generate_id_png_map():
    url = "https://raw.githubusercontent.com/DGP-Studio/Snap.Metadata/main/Genshin/CHS/Avatar.json"
    response = httpx.get(url).json()
    data = {record["Id"]: "./AvatarIcon/"+record["Icon"]+".png" for record in response}
    return data
