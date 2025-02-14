import requests, json, os, time, re
from lxml import etree
from PIL import Image
import xlwings as xw

def enfloat(l: list):
    for i in l:
        try:
            yield float(i)
        except ValueError:
            yield i

def extractDemand(re: dict):
    info, img, nzx, cncy, cwxr, clfz = {}, {}, {}, {}, {}, {}
    info['id'], info['model_name'], info['level'], info['model_number'], info['company'] = \
        re['id'], re['brand_title'] + re['car_type_title'], re['level_title'], re['model_number'], re['company']
    nzx['general'], nzx['structural'], nzx['repairing'], nzx['economics'], nzx['compatibility'] = \
        re['nzx'], re['nzx_1'], re['nzx_2'], re['nzx_3'], re['nzx_4']
    cncy['general'], cncy['driver_side_g'], cncy['side_impact_g'], cncy['roof_g'], cncy['seat_g'] =\
        re['cncy'], re['cncy_1'], re['cncy_2'], re['cncy_3'], re['cncy_4']
    cwxr['general'], cwxr['pedestrian_protection'] = re['cwxr'], re['cwxr_1']
    clfz['general'], clfz['safety_assist'] = re['clfz'], re['clfz_1']
    info['nzx'], info['cncy'], info['cwxr'], info['clfz'] = nzx, cncy, cwxr, clfz
    img['banner'], img['logo'] = re['image'], re['img']
    return info, img


def fetchInfo(info: dict, xpaths: list, imgurls: dict):
    url = "https://ciasi.org.cn/resultDetail?id=" + str(info['id'])
    html_page = etree.HTML(requests.get(url).text)
    # html_page = etree.HTML(requests.get(url, verify=False).text) # in case of certification expiring, use this line instead

    # fetching configurations
    configurations = html_page.xpath(xpaths['configurations'])[1:]
    configs = []
    for config_class in configurations:
        config_item = html_page.xpath(f"""//*[@class="{config_class}"]/div/div[@class="pur_le_item"]/div[@class="pur_l_txt"]/p/text()""")
        config_status = html_page.xpath(f"""//*[@class="{config_class}"]/div/div[@class="pur_le_item"]/div[@class="pur_l_rig"]/div/img/@src""")
        config_status_translate = {"greendi": "标配", "radca": "未配备", "yellowqu": "选配"}
        config_status = [config_status_translate[i.split("-")[-1].split(".")[0]] for i in config_status]
        config = dict(zip(config_item, config_status))
        config['name'] = html_page.xpath(f"""//*[@class="{config_class}"]/div/div[1]/p/text()""")[0]
        configs.append(config)
    info['configs'] = configs

    # anti-crash
    info['nzx']['structural_score'] = float(html_page.xpath(xpaths['structural_score'])[0])
    info['nzx']['repairing_score'] = float(html_page.xpath(xpaths['repairing_score'])[0])
    info['nzx']['economics_score'] = float(html_page.xpath(xpaths['economics_score'])[0])
    info['nzx']['collision_beam_size'] = html_page.xpath(xpaths['collision_beam_size'])
    info['nzx']['airbag_status'] = "未起爆" if html_page.xpath(xpaths['airbag_status'])[0].split("-")[-1].split(".")[0] == "radca" else "起爆"
    info['nzx']['price'] = float(re.sub('[^\u0030-\u0039\u002e]', '', html_page.xpath(xpaths['price'])[-1].strip()))
    imgurls['nzx'] = html_page.xpath(xpaths['nzx_img'])

    # passenger protection
    driverside_detail = html_page.xpath(xpaths['driver_side_25percent'])
    driverside_detail.append(html_page.xpath(xpaths['structure_observation'])[0])
    driverside_detail.append(html_page.xpath(xpaths['fuel_electricity'])[0])
    info['cncy']['driver_side'] = driverside_detail
    imgurls['driver_side'] = html_page.xpath(xpaths['driver_side_25percent_image'])
    passengerside_detail = html_page.xpath(xpaths['passenger_side_25percent_general'])
    passengerside_detail.extend(html_page.xpath(xpaths['passenger_side_25percent']))
    passengerside_detail.extend(html_page.xpath(xpaths['passenger_side_observation']))
    passengerside_detail.extend(html_page.xpath(xpaths['passenger_side_25percent_occupant']))
    info['cncy']['passenger_side'] = passengerside_detail
    imgurls['passenger_side'] = html_page.xpath(xpaths['passenger_side_25percent_image'])
    side_impact_detail = html_page.xpath(xpaths['side_impact1'])
    side_impact_detail.extend(html_page.xpath(xpaths['side_impact2']))
    info['cncy']['side_impact'] = side_impact_detail
    imgurls['side_impact'] = html_page.xpath(xpaths['side_impact_image'])
    roof_score = [float(i) for i in html_page.xpath(xpaths['roof'])]
    info['cncy']['roof'] = roof_score
    imgurls['roof'] = html_page.xpath(xpaths['roof_image'])
    seat_score = [float(i) for i in html_page.xpath(xpaths['seat'])]
    seat_score.extend(html_page.xpath(xpaths['static_dynamic']))
    seat_manufacturer = [i.split(":")[-1].strip() for i in html_page.xpath(xpaths['seat_manufacturer'])[0].strip().split(",")]
    info['cncy']['seat'] = seat_score
    info['cncy']['seat_manufacturer'] = seat_manufacturer

    # pedestrian protection
    # info['cwxr']["score"] = [float(i) for i in html_page.xpath(xpaths['pedestrian'])[:3]]
    info['cwxr']["score"] = list(enfloat((html_page.xpath(xpaths['pedestrian'])[:3])))

    # safety assist
    safety_assist = list(enfloat(html_page.xpath(xpaths['safety_assist'])[:5]))
    safety_assist.append(html_page.xpath(xpaths['sensor'])[0][6:].split("，")[0].split("。"))
    info['clfz']['score'] = safety_assist
    imgurls['safety_assist'] = html_page.xpath(xpaths['safety_assist_image'])
    return info, imgurls

def imageCrawl(urls: dict):
    for item in urls:
        if type(urls[item]) == str:
            if urls[item].startswith("https:"):
                url = urls[item]
            else:
                url = "https://ciasi.org.cn/upload/" + urls[item]
            with open(f"assets/{item}.{url.split('.')[-1]}", 'wb') as imgfile:
                imgfile.write(requests.get(url).content)
                # imgfile.write(requests.get(url, verify=False).content) # in case of certification expiring, use this line instead
        elif type(urls[item]) == list:
            for i in range(len(urls[item])):
                if urls[item][i].startswith("https:"):
                    url = urls[item][i]
                else:
                    url = "https://ciasi.org.cn/" + urls[item][i]
                with open(f"assets/{item + str(i)}.{url.split('.')[-1]}", 'wb') as imgfile:
                    imgfile.write(requests.get(url).content)
                    # imgfile.write(requests.get(url, verify=False).content) # in case of certification expiring, use this line instead
    return 1


def insertImageSubfuncion(sht:xw.sheets, img_path: str, position: str): # type: ignore
    cell = sht.range(position)
    imgw, imgh = Image.open(img_path).size
    cellh, cellw = cell.height, cell.width
    scale = cellh/imgh if cellh/imgh < cellw/imgw else cellw/imgw
    scale = int(scale*100)/100
    img_width, img_height = imgw*scale, imgh*scale
    left = cell.left + (cellw - img_width) / 2
    top = cell.top + (cell.height - img_height) / 2
    sht.pictures.add(os.path.join(os.getcwd(),img_path), left=left, top=top, width=img_width, height=img_height)
    time.sleep(0.2) # in case of pywintypes.com_error: (-2147352567...). no idea why this error occurs.

def insertImage(sht, pos):
    rel_path = "assets"
    for item in os.listdir(rel_path):
        img_file = os.path.join(rel_path,item)
        if item.split(".")[0] in ['logo', 'banner']:
            insertImageSubfuncion(sht, img_file, pos[item.split(".")[0]])
        else:
            insertImageSubfuncion(sht, img_file, pos[item.split(".")[0][:-1]][int(item.split(".")[0][-1])])
        time.sleep(0.3)
    return sht

def insertInfo(sht, pos, info):
    for item in info:
        if item == "configs":
            for i in range(len(info[item])):
                for subitem in info[item][i]:
                    sht.range(pos[item][subitem][i]).value = info[item][i][subitem]
        elif type(info[item]) == dict:
            subitem = info[item]
            for info_item in subitem:
                if type(subitem[info_item]) == str or type(subitem[info_item]) == float:
                    sht.range(pos[item][info_item]).value = subitem[info_item]
                elif type(subitem[info_item]) == list:
                    for i in range(len(subitem[info_item])):
                        sht.range(pos[item][info_item][i]).value = subitem[info_item][i]
        elif item == "id":
            pass # not needed in charts, only when fetching informations
        else:
            sht.range(pos[item]).value = info[item]
    return sht

def checkFolder(folder: str) -> None:
    if os.path.isdir(folder):
        pass
    else:
        os.mkdir(folder)
    for files in os.listdir(folder):
        os.unlink(os.path.join(folder, files))

with open("json/elements.json", 'r') as f:
    xpaths = json.load(f)
with open("json/pos.json", 'r') as f2:
    positions = json.load(f2)
img_pos, info_pos = positions['img_pos'], positions['info_pos']
results = "https://ciasi.org.cn/getResultList?p_id=&c_id=&gui_id%5B%5D=464"
info = json.loads(requests.get(results).text)["data"]
# info = json.loads(requests.get(results, verify=False).text)["data"]  # in case of certification expiring, use this line instead
xw.App(visible=False)
wb = xw.Book("files/template.xlsx")
template = wb.sheets['template']
for testing_year in info["year_list"]:
    year = testing_year["year"]
    title = f"CIASI {year}年测试车辆指数"
    models = testing_year["list"]
    for model in models:
        checkFolder("assets")
        info, img_urls = extractDemand(model)
        info, img_urls = fetchInfo(info, xpaths, img_urls)
        try:
            sht = template.copy(after=wb.sheets[year], name=info["model_name"])
        except ValueError:
            print(f"{info['model_name']} already exists")
            continue
        print(f"sheet {info['model_name']} created")
        status = imageCrawl(img_urls)
        sht.range("F1").value = title
        sht = insertInfo(sht, info_pos, info)
        sht = insertImage(sht, img_pos)
wb.save()
wb.close()