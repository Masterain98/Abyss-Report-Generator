import httpx
import metadata

id_to_png_map = metadata.generate_id_png_map()


def get_overview():
    url = "https://homa.snapgenshin.com/Statistics/Overview"
    response = httpx.get(url).json()["data"]
    return response


def get_floor_12_top_20_attendance_rate():
    url = "https://homa.snapgenshin.com/Statistics/Avatar/AttendanceRate"
    response = httpx.get(url).json()["data"][3]["Ranks"]
    data = {record["Item"]: record["Rate"] for record in response}
    data = {k: v for k, v in sorted(data.items(), key=lambda item: item[1], reverse=True)[:20]}
    data = {id_to_png_map[k]: f"{v * 100:.2f}%" for k, v in data.items()}
    return data


def get_floor_12_top_20_utilization_rate():
    url = "https://homa.snapgenshin.com/Statistics/Avatar/UtilizationRate"
    response = httpx.get(url).json()["data"][3]["Ranks"]
    data = {record["Item"]: record["Rate"] for record in response}
    data = {k: v for k, v in sorted(data.items(), key=lambda item: item[1], reverse=True)[:20]}
    data = {id_to_png_map[k]: f"{v * 100:.2f}%" for k, v in data.items()}
    return data


def get_floor_12_top_3_team_combination():
    url = "https://homa.snapgenshin.com/Statistics/Team/Combination"
    response = httpx.get(url).json()["data"][3]
    up_data = response["Up"]
    up_data = {record["Item"]: record["Rate"] for record in up_data}
    up_data = {k: v for k, v in sorted(up_data.items(), key=lambda item: item[1], reverse=True)[:3]}
    up_data = [{"uxxx-1-png": id_to_png_map[int(k.split(",")[0])],
                "uxxx-2-png": id_to_png_map[int(k.split(",")[1])],
                "uxxx-3-png": id_to_png_map[int(k.split(",")[2])],
                "uxxx-4-png": id_to_png_map[int(k.split(",")[3])],
                "count": v} for k, v in up_data.items()]
    count = 1
    for this_dict in up_data:
        for item in ["uxxx-1-png", "uxxx-2-png", "uxxx-3-png", "uxxx-4-png"]:
            this_dict[item.replace("uxxx", f"u{count}")] = this_dict.pop(item)
        count += 1

    down_data = response["Down"]
    down_data = {record["Item"]: record["Rate"] for record in down_data}
    down_data = {k: v for k, v in sorted(down_data.items(), key=lambda item: item[1], reverse=True)[:3]}
    down_data = [{"dxxx-1-png": id_to_png_map[int(k.split(",")[0])],
                  "dxxx-2-png": id_to_png_map[int(k.split(",")[1])],
                  "dxxx-3-png": id_to_png_map[int(k.split(",")[2])],
                  "dxxx-4-png": id_to_png_map[int(k.split(",")[3])],
                  "count": v} for k, v in down_data.items()]
    count = 1
    for this_dict in down_data:
        for item in ["dxxx-1-png", "dxxx-2-png", "dxxx-3-png", "dxxx-4-png"]:
            this_dict[item.replace("dxxx", f"d{count}")] = this_dict.pop(item)
        count += 1

    return {"up": up_data, "down": down_data}
