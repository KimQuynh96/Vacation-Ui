from mimetypes import init
import re
from kq_hr_param import pr_ad
from kq_hr_login import driver,Keys
from kq_hr_lib_common import data,data,cm,time,WebDriverWait,By,EC,json,Keys
class xpath():
    def depart_has_user(i):
        depart_has_user = data["rq_vc"]["single_depart"] +str(i) + data["rq_vc"]["single_depart1"]
        return depart_has_user
    
    def is_user(i):
        is_user = data["rq_vc"]["sl_user"] + str(i) + data["rq_vc"]["sl_user1"]
        return is_user

    def cc_name(i):
        cc_name = data["rq_vc"]["user_name_cc"] + str(i) + data["rq_vc"]["user_name_cc1"]
        return cc_name
    
    def select_cc(i):
        select_cc = data["rq_vc"]["bt_cc_cc1"]+str(i)+data["rq_vc"]["bt_cc_cc2"]
        return select_cc

    def vc_type(j):
        vc_type = cm.xpath2(pr_ad.dt_ad["vc_type"],str(j),pr_ad.dt_ad["vc_type1"])
        return vc_type

    def manager(type,i):
        if type == "ma" :
            return cm.xpath("tr",str(i),pr_ad.dt_ad["manager"])
        elif type == "ap" :
            return cm.xpath("tr",str(i),pr_ad.dt_ad["approve"])
        elif type == "ad":
            return cm.xpath("tr",str(i),pr_ad.dt_ad["adjust"])
        else :
            return cm.xpath("tr",str(i),pr_ad.dt_ad["settlement"])

    def vc_name_delete(i):
        vc_name_delete = cm.xpath2(pr_ad.dt_ad["vc_name"],str(i),pr_ad.dt_ad["vc_name1"])
        return vc_name_delete

    def vc_name(i):
        vc_name = cm.xpath2(pr_ad.dt_ad["vc_name"],str(i),pr_ad.dt_ad["vc_name1"])
        return vc_name

    def search():
        search = cm.xpath2(pr_ad.dt_ad["vc_name"],str(j),pr_ad.dt_ad["vc_name1"])
        return search

    def manager_name(i):
        manager_name = cm.xpath("tr",str(i),pr_ad.dt_ad["manager"])
        return manager_name

    def ic_edit(manager_before_modify):
        ic_edit = cm.xpath2(pr_ad.dt_ad["ic_edit"],str(manager_before_modify["no"]),pr_ad.dt_ad["ic_edit1"])
        return ic_edit

    def ic_delete(manager_before_delete):
        ic_delete = cm.xpath2(pr_ad.dt_ad["ic_delete"],str(manager_before_delete["no"]),pr_ad.dt_ad["ic_delete1"])
        return ic_delete

    def ardec_name(i):
        ardec_name = cm.xpath2(pr_ad.dt_ad["ardec0"],str(i),pr_ad.dt_ad["ardec"])
        return ardec_name
    
    def delete_ardec(i):
        delete_ardec = cm.xpath2(pr_ad.dt_ad["ardec0"],str(i),pr_ad.dt_ad["ardec"])
        return delete_ardec
    def ax_name(i):
        ax_name = cm.xpath2(pr_ad.dt_ad["apex"],str(i),pr_ad.dt_ad["apex1"])
        return ax_name
    

    def open_filter(select_status):
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR,pr_ad.dt_ad["open_filter"]))).click()
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH,pr_ad.dt_ad[select_status]))).click()

    def filter_text(type , sta_name_to_filter):
        if type =="p":
            return "-Filter by type is "+ sta_name_to_filter + "<Pass>" 
        else :
            return "-Filter by type is "+ sta_name_to_filter + "<Fail>"
    
    def par_manager():
        list_manager = {
            "no":"",
            "name":"",
            "approve":"",
            "adjust":"",
            "settlement":""
            }
        return list_manager