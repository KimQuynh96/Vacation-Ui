

from kq_hr_lib_common import cm,driver,data
from kq_hr_lib_common import datetime,json,re,time,pipe,EC,By,WebDriverWait,Keys,relativedelta
from kq_hr_param import pr_rq ,pr_mn,submenu_my_vacation
param = json.loads(submenu_my_vacation())

class xpath():
    def depart_has_user(i):
        depart_has_user = pr_rq.rq_vc["single_depart"] + str(i) + pr_rq.rq_vc["single_depart1"]
        return depart_has_user

    def click_on_depart(i):
        return pr_rq.rq_vc["single_depart"] + str(i) + pr_rq.rq_vc["single_depart1"]
    
    def is_user(j) :
        return pr_rq.rq_vc["sl_user"]+ str(j) + pr_rq.rq_vc["sl_user1"]

    def select_user(i,j):
        data1 = pr_rq.rq_vc["depart"] + str(i) + pr_rq.rq_vc["depart_on"]
        data2 = pr_rq.rq_vc["user_name_cc"] + str(j) + pr_rq.rq_vc["user_name_cc1"]
        return data1 +  data2

    def click_user(i,j):
        data1 = pr_rq.rq_vc["depart"] + str(i) + pr_rq.rq_vc["depart_on"]
        data2 = pr_rq.rq_vc["bt_cc_cc1"] + str(j) + pr_rq.rq_vc["bt_cc_cc2"]
        return data1 + data2

    def get_date(i,j):
        return data["date_cu_tc"] + str(i) + "]/div[" + str(j) +"]"

    def vc_start(vacation_name):
        return vacation_name[vacation_name.rfind("\n") + 1 : int(vacation_name.rfind("\n")) +11]
    
    def vc_change(vacation, vacation_name):
        return vacation_name[0:int(vacation_name.rfind("\n"))] + "["+vacation["start"] + " ~ " + vacation["expiration"] + "]"
    
    def us_change(vacation_name):
        return vacation_name[None:int(vacation_name.rfind("("))-1] + vacation_name[int(vacation_name.rfind(")"))+2: None]
    
    def view_detail(i):
        return cm.xpath3("tr" , str(i), pr_rq.rq_vc["ic_detail"],i-1, pr_rq.rq_vc["ic_detail1"])

    def vc_available(i,type):
        if type   == "to":
            return pr_rq.rq_vc["tr"]+str(i)+"]/td[2]"
        elif type == "na" :
            return pr_rq.rq_vc["tr"]+str(i)+"]/td[1]"
        elif type == "us":
            return pr_rq.rq_vc["tr"]+str(i)+"]/td[3]"
        elif type == "re" :
            return pr_rq.rq_vc["tr"]+str(i)+"]/td[4]" 
        elif type == "ex" :
            return pr_rq.rq_vc["tr"]+str(i)+"]/td[5]"

    def vc_use_name(vacation_name):
        return vacation_name[None:int(vacation_name.rfind("("))-1] + vacation_name[int(vacation_name.rfind(")"))+2: None]

    def vc_use_days(vacation_name):
        return vacation_name[int(vacation_name.rfind("(") + 2) : int(vacation_name.rfind(")"))]

    def vc_use_number_days(vacation_name, type):
        if type == "d" :
            return vacation_name[int(vacation_name.rfind("(")) + 1: int(vacation_name.rfind("D"))]
        else :
            return vacation_name[int(vacation_name.rfind("(")) + 1: int(vacation_name.rfind("H"))]

    def vc_use_number_hours(vacation_name):
        return vacation_name[int(vacation_name.rfind("D")) + 1 : int(vacation_name.rfind("H"))]
    
    def approver_name(i):
        return pr_rq.rq_vc["approver"] + str(i) + pr_rq.rq_vc["approver_one"]
    
    def tc_hour_use(hour_use,type):
        if type == "d" :
            return hour_use[int(hour_use.rfind("Use:")): int(hour_use.rfind("] ["))]
        else :
            return hour_use[int(hour_use.rfind("Real Used:") + 10) : int(hour_use.rfind("H )"))]
    
    def tc_hour(hour_use):
        return int(re.search(r'\d+',hour_use).group(0)) 
    
    def sd_date(i):
        return cm.xpath("tr",str(i),pr_rq.rq_vc["request"])

    def sd_status(i):
        return cm.xpath("tr",str(i),pr_rq.rq_vc["re_status"])
    
    def total_remain(i):
        return pr_rq.rq_vc["tr"]+str(i)+"]/td[4]"

    def sl_approver(i):
        return pr_rq.rq_vc["approver_name3"]+str(i)+ pr_rq.rq_vc["approver_name4"]

    def sl_name(i):
        return pr_rq.rq_vc["ap_name"]+str(i)+ pr_rq.rq_vc["ap_name_two"]

    def cc_name(i):
        return pr_rq.rq_vc["cc_name1"]+"div["+str(i)+ pr_rq.rq_vc["cc_name2"]

    def click_date(i,j):
        return "//tr[" + str(i) + "]/td[" + str(j) + "]/span"
    
    def select_date(selected_date,type):
        if type == "y" :
            return selected_date[None: int(selected_date.rfind("["))].replace(" ", "")
        else :
            return selected_date.replace(" ", "")

    def vr_replace(vacation_request,type):
        if type == "da" :
            return vacation_request["vc_date"].replace("\n", "")
        else :
            return vacation_request["vc_name"].replace("\n", "").replace(" ", "")

    def data_column(data_column,type):
        if type == "d" :
            return data_column[None : int(data_column.rfind("D"))]
        else :
            return data_column[int(data_column.rfind("H")) -1: int(data_column.rfind("H"))]

    def total_hour(tp1,tp2,oneday):
        return int(tp1["hour"]) + int(tp2["hour"]) + int(tp1["day"]*oneday) + int(tp2["day"]*oneday)

    def li_request(i,type):
        if type == "na" :
            return cm.xpath("tr",str(i), pr_rq.rq_vc["re_name"])
        elif type == "da":
            return cm.xpath("tr",str(i), pr_rq.rq_vc["re_vc_date"])
        elif type == "us" :
            return cm.xpath( pr_rq.rq_vc["tr_re_use"],str(i), pr_rq.rq_vc["re_use"])
        elif type == "rd" :
            return cm.xpath("tr",str(i), pr_rq.rq_vc["re_date"])
        elif type == "st" :
            return cm.xpath("tr",str(i), pr_rq.rq_vc["re_status"])

    def in_detail(info_vc,type):
        if type == "if":
            return info_vc["vc_date"].rfind("All day Off-All day Off")
        elif type == "re" :
            return "Vacation Date : "+ info_vc["vc_date"].replace("All day Off-All day Off", "")
        elif type == "no" :
            return "Vacation Date : "+ info_vc["vc_date"]
        elif type == "dt" :
            return "Request Date : " + info_vc["request_date"]
        elif type == "us" :
            return "Used : "+ info_vc["used"]
        else :
            return "Reason : "+ info_vc["reason"]

    def vacation_date(type,vc_date):
        if type == "d" :
            return vc_date.replace("(", "").replace(")", "")
        else :
            return vc_date[None: int(vc_date.rfind("H"))-1] + vc_date[int(vc_date.rfind("H")): len(vc_date)-1] 

    def send_request():
        return cm.msg("n" ,"Send Request")
    
    def today():
        return str(datetime.date.today()).replace(" ","")

    def filter(status,status_name):
        if status == "p" :
            return  "  +Filter by status is "+ status_name + " <Pass>"
        else :
            return  "  +Filter by status is "+ status_name + " <Fail>"

    def filter_name_my(j):
        return cm.xpath("tr",str(j),pr_rq.rq_vc["re_status"])

    def filter_name_cc(j):
        return cm.xpath("tr",str(j),pr_rq.rq_vc["re_status_cc1"])

    def filter_page(total_ic):
        return pr_mn.mn_pro["ic_next_page"]+ str(len(total_ic)-1)+pr_mn.mn_pro["ic_next_page1"]

    def par_vacation():
        vacation   = {
            "total"        :"",
            "used"         :"",
            "remain"       :"",
            "start"        :"",
            "expiration"   :"",
            "vacation_name":""
            }
        return vacation
    
    def par_select_approver():
        select_approver = {
            "result_approver"   : False ,
            "approver_name"     : True  ,
            "approval_line"     : False ,
            "approval_exception": False 
            }
        return select_approver
    
    def par_usage_settings():
        usage_settings = {
            "vacation_name"  : "",
            "number_of_days" : "",
            "number_of_hours": "",
            "use_hour_unit"  : "",
            "use_half_day"   : "",
            "hour_use"       : ""
            }
        return usage_settings
    
    def par_info():
        info = {
            "vc_date"      : "",
            "used"         : "",
            "request_date" : "",
            "approver"     : "",
            "reason"       : ""
            }
        return info
    
    def par_vc_rq():
        vc_rq = {
            "vc_name"     : "",
            "vc_date"     : "",
            "request_date": "",
            "status"      : "Request"
            }
        return vc_rq
    
    def par_infor_request():
        infor_request = {
            "no"           :"",
            "vacation_name":"",
            "vacation_date":"",
            "use"          :"",
            "request_date" :"",
            "status"       :"",
            "icon_cancel"  :"",
            "vacation_time":""
            }
        return infor_request
    
    def par_approver_result():
        result = { 
            "rs_resaon"  :False ,
            "rs_approver":False ,
            "info_after" :""    ,
            "info_before":""    }
        return result
    
    def par_vacation_request(request_date):
        vacation_request = {
            "vc_name"     : ""           ,
            "vc_date"     : ""           ,
            "request_date": request_date ,
            "status"      : "Request"
            }
        return vacation_request

    def par_choose_vacation(hu_use,rs_sl_vc,us_hu_un,vc_re,vc_na):
        result = {
            "hour_use"         :hu_use,
            "result_select_vc" :rs_sl_vc,
            "use_hour_unit"    :us_hu_un,
            "vacation_request" :vc_re,
            "vc_name"          :vc_na
            
        }
        return result

    def par_choose_date(rs_date,vc_date,vc_rq):
        result = {
        "result_date"         :rs_date,
        "vc_date"             :vc_date,
        "vacation_request"    :vc_rq,
        }
        return result
    
    def par_all_day():

        result = {
        "result_date"          :"",
        "vc_date"              :"",
        "oneday"               :"",
        "number_before_request":"",
        "result_select_vc"     :"",
        "use_hour_unit"        :"",
        "hour_use"             :"",
        "vc_name"              :""
        }
        return result

    def par_number_of_days(before,after,type):
        result = {
            "before"       : before.number_before_request ,
            "after"        : after,
            "hour_use"     : before.hour_use ,
            "vc_name"      : before.vc_name,
            "oneday"       : before.oneday,
            "use_hour_unit": before.use_hour_unit,
            "type_request" : type
        }
        return result
    
    def par_request(list_status_before):
        result = {
            "sta_name_to_filter" : "Request" ,
            "select_status"      : "select_st_request",
            "submenu"            : "myvacation" ,
            "list_status_before" : list_status_before 
            
        }
        return result
    
    def par_progress(list_status_before):
        result = {
            "sta_name_to_filter" : "Progressing" ,
            "select_status"      : "select_st_progressing",
            "submenu"            : "myvacation" ,
            "list_status_before" : list_status_before 
            
        }
        return result
    
    def par_completed(list_status_before):
        result = {
            "sta_name_to_filter" : "Completed" ,
            "select_status"      : "select_st_completed",
            "submenu"            : "myvacation" ,
            "list_status_before" : list_status_before 
            
        }
        return result

    def par_date_and_sta(vc_date,vc_stat):
        vaca = {
            "vc_date" : vc_date ,
            "status"  : vc_stat ,
        }
        return vaca

    
    def cc_request(list_status_before):
        result = {
            "sta_name_to_filter" : "Request" ,
            "select_status"      : "select_st_request",
            "submenu"            : "viewcc" ,
            "list_status_before" : list_status_before 
            
        }
        return result

    def cc_progress(list_status_before):
        result = {
            "sta_name_to_filter" : "Progressing" ,
            "select_status"      : "select_st_progressing",
            "submenu"            : "viewcc" ,
            "list_status_before" : list_status_before 
            
        }
        return result

    def cc_completed(list_status_before):
        result = {
            "sta_name_to_filter" : "Completed" ,
            "select_status"      : "select_st_completed",
            "submenu"            : "viewcc" ,
            "list_status_before" : list_status_before 
            
        }
        return result

class type_vc():
    def check_type(type_request , text):

        if type_request  == "all" :
            text_request = "  *All day : "
        elif type_request == "hour" :
            text_request  = "  *Hour Unit : "
        elif type_request == "vc_con" :
            text_request  = "  *Vacation consecutive : "
        elif type_request == "half_day" :
            text_request  = "  *Half day : "
        return text_request + text

    def hour_use(type_request):
        if type_request == "all" :
            hour_use = "1D"
        elif type_request == "half_day" :
            hour_use = "4H"
        elif type_request == "hour" :
            hour_use = "2H"
        else :
            hour_use = "2D"
        return hour_use

class cm_fu():

    def my_vc():
        cm.msg("n","II.MY VACATION")
        driver.find_element_by_link_text("My Vacation Status").click()
        driver.implicitly_wait(5)

    def click_on_button_to_request():
        driver.find_element_by_xpath( pr_rq.rq_vc["bt_request_be"]).click()
        driver.find_element_by_css_selector( pr_rq.rq_vc["bt_request_af"]).click()

    def approver_name(total_app,list_approver):
        try :
            i = 1
            while i <= total_app :
                app_name = driver.find_element_by_xpath(xpath.sl_name(i)).text
                app_name = app_name.replace(" ", "")
                list_approver.append(app_name)
                i = i +1  
        except :
            pass
        return list_approver


    def check_reason(info_before,info_after,result,type_request):
        reason = driver.find_element_by_xpath(pr_rq.rq_vc["content_reason"]).text
        info_after["reason"]  = reason
        info_before["reason"] =  pr_rq.rq_vc["reason_text"]
       
        if info_after["reason"] == info_before["reason"] :
            cm.msg("p" ,type_vc.check_type(type_request , pr_rq.msg_re["pass_detail_reason"]))
            result["rs_resaon"] = True 
        else:
            fail_reason = type_vc.check_type(type_request , pr_rq.msg_re["fail_detail_reason"])
            cm.xlsx( pr_rq.re_detail["fail"] , fail_reason )

    def vacation_request(vacation_request,i):
        vacation_request["vc_name"]      = driver.find_element_by_xpath(xpath.li_request(i,"na")).text
        vacation_request["vc_date"]      = driver.find_element_by_xpath(xpath.li_request(i,"da")).text
        vacation_request["request_date"] = driver.find_element_by_xpath(xpath.li_request(i,"rd")).text
        vacation_request["status"]       = driver.find_element_by_xpath(xpath.li_request(i,"st")).text
        vacation_request["vc_date"]      = xpath.vr_replace(vacation_request,"da")
        vacation_request["vc_name"]      = xpath.vr_replace(vacation_request,"na")
        return vacation_request

    def infor(vacation,title,type_request): 
        hour_use      = type_vc.hour_use(type_request)
        vacation_name = "Vacation Name : "+ vacation["vacation_name"]
        total         = "Total : " + vacation["total"]
        used          = "Used : " + vacation["used"]
        remain        = "Remain : " + vacation["remain"]
        hour          = "Hour use for request : " + str(hour_use)
        data          = "  +" +title +" [ "+vacation_name +" | " + total 
        data1         = " | " + used + " | " + remain +" | " + hour + " ] "
        return data + data1

    def result_number(number_before,number_after,msg_list,excel,result,type_request):
        if number_before != number_after:
            cm.xlsx(excel["fail"],type_vc.check_type(type_request ,msg_list["fail"]))
            result = False
        else:
            cm.msg("p",type_vc.check_type(type_request ,msg_list["pass"]))
        return result

    def result(number_before,number_after,msg_list,excel,type_request):
        if number_before == number_after:
            cm.msg("p",type_vc.check_type(type_request ,msg_list["pass"]))
        else:
            cm.xlsx(excel["fail"],type_vc.check_type(type_request ,msg_list["fail"]))

    def result_all(a,b,c,msg_list,excel,type_request):
        if a ==  True and b == True and c == True :
            cm.xlsx(excel["pass"],type_vc.check_type(type_request ,msg_list["pass"]))
        else:
            cm.xlsx(excel["fail"],type_vc.check_type(type_request ,msg_list["fail"]))

    def infor_detail(title,info_vc):
        result_approver = isinstance(info_vc["approver"], str)
        if result_approver == False :
            approver = ""
            for name in info_vc["approver"] :
                approver = approver  + name + ","
        else :
            approver =  info_vc["approver"]

        date = xpath.in_detail(info_vc,"if")
        if date > 0 :
            vacation_date = xpath.in_detail(info_vc,"re")
        else :
            vacation_date = xpath.in_detail(info_vc,"no")

        request_date = xpath.in_detail(info_vc,"dt")
        used         = xpath.in_detail(info_vc,"us")
        reason       = xpath.in_detail(info_vc,"ld")
        approver     = "Approver : " + approver
        data         = "  +" + title +" [ " + vacation_date + " | " + request_date 
        data1        = " | " + used + " | " + approver + " | " + reason + " ] "
        cm.msg("t" , data + data1 )

    def click_on_request_button():
        driver.find_element_by_xpath(data["rq_vc"]["bt_request_be"]).click()
        driver.find_element_by_css_selector(data["rq_vc"]["bt_request_af"]).click()

    def info_cc(selected_cc , saved_cc):
        print("  +Info selected cc :" , selected_cc)
        print("  +Info save cc     :" , saved_cc)
    
    def next_date(request_date):
        
        if int(data["month"][str(request_date.month)]) == request_date.day:
            request_date = request_date + relativedelta(month=request_date.month+1) + relativedelta(day=1)
    
        elif request_date.weekday() == 5 :
            request_date = request_date + relativedelta(day=request_date.day +2)
           
        elif request_date.weekday() == 6 :
            request_date = request_date + relativedelta(day = request_date.day +1)
            if int(data["month"][str(request_date.month)]) == request_date.day:
                request_date= request_date + relativedelta(month=request_date.month+1) + relativedelta(day=1)
        else :
            request_date = request_date + relativedelta(day=request_date.day +1)

        return request_date

    def split_date_from_continuous_date(continuous_date,date_used) :
       
        if continuous_date.rfind("~") > 0 :
            start_date  = continuous_date[None: int(continuous_date.rfind("~"))]
            start_date  = datetime.datetime.strptime(start_date , '%Y-%m-%d').date()
            end_date    = continuous_date[int(continuous_date.rfind("~"))+1: None]
            end_date    = datetime.datetime.strptime(end_date , '%Y-%m-%d').date()
            next_date_1 = start_date
            while next_date_1 !=  end_date :
                date_used.append(str(next_date_1))
                if start_date == end_date :
                    break
                next_date_1 = cm_fu.next_date(next_date_1)
            date_used.append(str(end_date))
        else :
            date_used.append(continuous_date)
       
    def get_vacation_date(continuous_date):
        date_used = []
        if continuous_date.rfind("~") > 0 :
            start_date  = continuous_date[None: int(continuous_date.rfind("~"))]
            start_date  = datetime.datetime.strptime(start_date , '%Y-%m-%d').date()
            end_date    = continuous_date[int(continuous_date.rfind("~"))+1: None]
            end_date    = datetime.datetime.strptime(end_date , '%Y-%m-%d').date()
            next_date_1 = start_date
            while next_date_1 !=  end_date :
                date_used.append(str(next_date_1))
                if start_date == end_date :
                    break
                next_date_1 = cm_fu.next_date(next_date_1)
            date_used.append(str(end_date))
        else :
            date_used.append(continuous_date)        
        return date_used


    def choose_end_date(request_date,date_used):
        start_date   = request_date
        request_date = cm_fu.next_date(request_date)
        
        if  start_date != request_date :
            if request_date.weekday() == 5 :
                request_date = request_date + relativedelta(day = request_date.day +2)
                if int(data["month"][str(request_date.month)]) == request_date.day:
                    request_date = request_date + relativedelta(month = request_date.month + 1) + relativedelta(day = 1)

            if  request_date.weekday() == 6 :
                request_date = request_date + relativedelta(day = request_date.day +1)
                if int(data["month"][str(request_date.month)]) == request_date.day:
                    request_date = request_date + relativedelta(month = request_date.month+1) + relativedelta(day = 1)

        end_date = request_date
        if str(end_date) not in date_used:
            return end_date
        else :
            return False

    def choose_start_date(date_used,request_date):
        # Find unused date , not saturday , not sunday , not holiday to use for request vacation # 
        if request_date.weekday() == 5 :
            request_date = request_date + relativedelta(day=request_date.day +2)
        if  request_date.weekday() == 6 :
            request_date= request_date + relativedelta(day=request_date.day +1)
        if str(request_date) in date_used  :
            request_date= request_date + relativedelta(day=request_date.day +1)
            if request_date.weekday() == 5 :
                request_date= request_date + relativedelta(day=request_date.day +2)
            if  request_date.weekday() == 6 :
                request_date= request_date + relativedelta(day=request_date.day +1)
            
            while str(request_date) in date_used  :
                request_date = cm_fu.next_date(request_date)
                if request_date.weekday() == 5 :
                    request_date = request_date + relativedelta(day=request_date.day +2)
                if  request_date.weekday() == 6 :
                    request_date= request_date + relativedelta(day=request_date.day +1)
        return request_date
    
    def click_date(request_date):
        if request_date.day < 25:
            for i in range(2,8) :
                for j in range(2,8):
                    date_at_calendar = driver.find_element_by_xpath(xpath.click_date(i,j))
                    if str(date_at_calendar.text) == str(request_date.day):
                        date_at_calendar.click()
                        return True
            return False
        else:  
            for i in range(2,8) :
                for j in range(2,8):
                    date_at_calendar = driver.find_element_by_xpath(xpath.click_date(i,j))
                    if str(date_at_calendar.text) == str(request_date.day):
                        date_at_calendar.click()
                        selected_date = driver.find_element_by_xpath(pr_rq.rq_vc["selected_date"]).text
                        if selected_date.rfind("[") > 0:
                            selected_date = xpath.select_date(selected_date,"y")
                        else:
                            selected_date = xpath.select_date(selected_date,"n")
                            #selected_date = selected_date[12: None].replace(" ", "")
                        if str(request_date) == selected_date:
                            return True
                        else:
                            date_at_calendar.click()   
                            return True
                           
            return False

    def get_days_and_hour(data_column):
        # Get days , hour of column data 4.5D , 4D4H , - #  
        number_day = {"day":"","hour":""}

        if data_column.replace(" ", "") == "-":
            number_day["day"] = float(0)
        elif data_column.rfind("D") < 0 :
            number_day["day"] = float(0)
        else :
            number_day["day"] = float(xpath.data_column(data_column,"d"))
            

        if data_column.rfind("H") < 0 :
            number_day["hour"] = float(0)
        else :
            number_day["hour"] = float(xpath.data_column(data_column,"h"))

        return number_day
    
    def change_hour_to_day(tp1,tp2,oneday,plus,hour_use,use_hour_unit,type_request):

        # USE HOUR UNIT FOR VACATION # 
        # The unit for calculation is hour ,convert to hour before calculation #
        if use_hour_unit == True :
            # Hour_use is int  ,ex hour_use = 4 #
            # Plus or minus data 2 column #
            if tp2 != None:
                tp1 = cm_fu.get_days_and_hour(tp1)
                tp2 = cm_fu.get_days_and_hour(tp2)
            
                if plus =="plus":
                    total_hour = xpath.total_hour(tp1,tp2,oneday)
                    day = total_hour // oneday
                    hour = total_hour % oneday
                
                if  plus == "minus":
                    l1 = int(tp1["day"])*oneday + int(tp1["hour"])
                    l2 = int(tp2["day"])*oneday + int(tp2["hour"])
                    total_hour_remain = l1-l2
                    if total_hour_remain < 0:
                        total_hour_remain = total_hour_remain*(-1)
                    day = total_hour_remain // oneday
                    hour = total_hour_remain % oneday
                
                if str(day) == "0" and str(hour) == "0":
                    return "0"
                else:
                    if str(day) == "0" :
                        return str(hour) + "H"
                    elif str(hour) == "0" :
                        return str(day) + "D"
                    else :
                        return str(day) + "D" + " " + str(hour) + "H"
            else:
            # Plus or minus data of 1 column with number #
                hour_use = int(hour_use)
                tp1 = cm_fu.get_days_and_hour(tp1)
            
                if plus =="plus":
                    if type_request == "half_day"  :
                        hour_use   = 4
                        total_hour = int(tp1["hour"])  + int(tp1["day"]*oneday) + hour_use

                    elif type_request  == "hour":
                        hour_use   = 2
                        total_hour = int(tp1["hour"])  + int(tp1["day"]*oneday) + hour_use

                    else:
                        total_hour = int(tp1["hour"])  + int(tp1["day"]*oneday) + hour_use*oneday
                    day  = total_hour // oneday
                    hour = total_hour % oneday

                
                if  plus =="minus":
                    l1 = int(tp1["day"])*oneday + int(tp1["hour"])
                    if type_request == "half_day"  :
                        hour_use          = 4
                        total_hour_remain = l1 - hour_use

                    elif type_request == "hour":
                        hour_use          = 2
                        total_hour_remain = l1 - hour_use

                    else:
                        total_hour_remain = l1 - int(hour_use)*oneday
                    day  = total_hour_remain // oneday
                    hour = total_hour_remain % oneday

                if str(day) == "0" and str(hour) == "0":
                    return "0"
                else:
                    
                    if str(day) == "0" :
                        return str(hour)+"H"
                    elif str(hour) == "0" :
                        return str(day)+"D"
                    else:
                        return str(day)+"D " + str(hour)+"H"
        else:
            # NOT USE HOUR UNIT FOR REQUEST #
            # The unit for calculation is days ,convert to days before calculation #
            # Hour used have to convert to day #
            
            if tp2 != None:
                tp1 = cm_fu.get_days_and_hour(tp1)
                tp2 = cm_fu.get_days_and_hour(tp2)
            
                # Day after plus #
                if plus == "plus":
                    day = float(tp1["day"]) + float(tp2["day"])
                
                # Day after minus #
                if  plus == "minus":
                    day = float(tp1["day"]) - float(tp2["day"])
                    if day < 0:
                        day = day*(-1)
                
                if str(day) != "0" :
                    if str(day)[int(str(day).rfind("."))+1: None] == "0":
                        return str(day)[None: int(str(day).rfind("."))]+"D"
                    else:
                        return str(day)+"D"
                else:
                    return "0"
            else:
                if type_request == "half_day" :
                    hour_use = 0.5

            # Hour_use is fload ,ex hour_use  =  0.5 #
            # Plus or minus data of 1 column with number #

                tp1 = cm_fu.get_days_and_hour(tp1)
                if plus =="plus":
                    day = float(tp1["day"]) + float(hour_use)
                
                if  plus == "minus":
                    day = float(tp1["day"]) - float(hour_use)
                    if day < 0 :
                        day = day *(-1)
            
                if str(day) != "0.0" :      
                    if str(day)[int(str(day).rfind("."))+1: None] == "0":
                        return str(day)[None: int(str(day).rfind("."))]+"D"
                    else:
                        return str(day)+"D"
                else :
                    return "0"
    
    def select_user_from_depart():
        list_department  = driver.find_elements_by_xpath(pr_rq.rq_vc["list_depart_cc"])    
        total_department = cm.total_data(list_department)
        for i in range(1,total_department):
            time.sleep(1)
            depart_has_user = cm.is_Displayed(xpath.depart_has_user(i)) 
            if depart_has_user == True :
                driver.find_element_by_xpath(xpath.click_on_depart(i)).click()
                total_user = driver.find_elements_by_xpath(pr_rq.rq_vc["list_user"])
                for j in range(1,len(total_user)+1):
                    is_user = cm.is_Displayed(xpath.is_user(j)) 
                    if is_user == True:
                        selected_cc_name = driver.find_element_by_xpath(xpath.select_user(i,j)).text
                        driver.find_element_by_xpath(xpath.click_user(i,j)).click()
                        return selected_cc_name      
        return False 

    def hours_set_from_time_card(type_request):

        # Specific working hours from time card  #  
        hour_use = driver.find_element_by_xpath(pr_rq.rq_vc["hour_use"]).text
        hour_use = xpath.tc_hour_use(hour_use,"d")

        if type_request == "all" :
            if len(hour_use) == 0 :
                return 8
            else :
                hour_use = xpath.tc_hour(hour_use)
                if hour_use == 1 :
                    return 8
                else :
                    return hour_use
        elif type_request == "hour":
            hour_use = driver.find_element_by_xpath(pr_rq.rq_vc["hour_use_h"]).text
            hour_use = xpath.tc_hour_use(hour_use,"h")
            return xpath.tc_hour(hour_use)
        else:
            return 4   
      
    def vacation_use_for_request():
        # Information about number of selected vacation for request vacation #  
        vacation = {"vacation_name":"","number_of_days":"","number_of_hours":""}
        vacation_name = driver.find_element_by_xpath(pr_rq.rq_vc["vacation_name"]).text
        if vacation_name.rfind("(") > 0 :
            vacation["vacation_name"] = xpath.vc_use_name(vacation_name)
        else :
            vacation["vacation_name"] = vacation_name

        days = xpath.vc_use_days(vacation_name)
        if int(days.rfind("D")) > 0 :
            vacation["number_of_days"] = xpath.vc_use_number_days(vacation_name,"d")
            if int(days.rfind("H")) < 0:
                vacation["number_of_hours"] = "0"
            else:
                vacation["number_of_hours"] = xpath.vc_use_number_hours(vacation_name)
        else:
            vacation["number_of_days"] = "0"
            if int(days.rfind("H")) > 0 :
                vacation["number_of_hours"] = xpath.vc_use_number_days(vacation_name,"h")
            else:
                vacation["number_of_hours"] = "0"

        if int(days.rfind("D")) < 0 and int(days.rfind("H")) < 0 :
            vacation["number_of_days"]  = "0"
            vacation["number_of_hours"] = "0"

        return vacation

    def hour_used(use_hour_unit,type_use):
        if bool(use_hour_unit) == True :
            if type_use =="all":
                hour_use = 1
            elif type_use == "hour":
                hour_use = cm_fu.hours_set_from_time_card(type_use)
            else:
                hour_use = 0.5
            return float(hour_use)
        else:
            if type_use == "all":
                hour_use = 1 
            else:
                hour_use = 4/10
            return float(hour_use)

    def click_date_time_card(request_date):
        day = request_date.day
        for i in range(1,6) :
            for j in range(1,6):
                if i == 1 :
                    date_to_click = driver.find_element_by_xpath(xpath.get_date(i,j))
                    date_at_calendar = date_to_click.text
                    if day > 25 and date_at_calendar == str(day) :
                        pass
                    else:
                        if date_at_calendar == str(day):
                            date_to_click.click()
                            return True
                else:
                    date_to_click = driver.find_element_by_xpath(xpath.get_date(i,j))
                    date_at_calendar = date_to_click.text
                    if date_at_calendar == str(day):
                        date_to_click.click()
                        return True
    
    def view_detail_used(type_request,use_hour_unit):
        
        if type_request == "all":
            return "1D"
        elif type_request == "vc_con":
            return "2D"
        elif type_request == "hour":
            return "2H"
        else:
            if use_hour_unit == True:
                return "4H"
            else:
                return "0.5D"

    def approver_list(check_no_approver):
        if len(check_no_approver) != 0 :
            return True
        else :
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_approver_no"])
            return False

    def select_approver():
        result ={
            "user_name"   : "",
            "is_selected" : ""
        }
        selected_user       = driver.find_element_by_xpath( pr_rq.rq_vc["sl_ap_firt"])
        result["user_name"] = selected_user.text
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, pr_rq.rq_vc["sl_ap_firt"]))).click()
        button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, pr_rq.rq_vc["bt_ap_firt"])))
        is_selected = button.is_selected()
        if is_selected == True :
            cm.msg("p" , pr_rq.msg_re["pass_approver_click"])
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, pr_rq.rq_vc["bt_add"]))).click()
            result["is_selected"] = True 
        else :
            result["is_selected"] = False
            cm.xlsx( pr_rq.re_vc["fail"] , pr_rq.msg_re["fail_approver_click"])
        return result

    def add_approver():
        if cm.is_Displayed( pr_rq.rq_vc["check_bt_save"]) == False :
            cm.msg("p" , pr_rq.msg_re["pass_approver_add"])
            driver.find_element_by_xpath( pr_rq.rq_vc["bt_save"]).click()
            return True
        else :
            cm.xlsx( pr_rq.re_vc["fail"] , pr_rq.msg_re["fail_approver_add"])
            return False

    def save_approver(user_name,select_approver):
        if cm.is_Displayed( pr_rq.rq_vc["bt_select_cc"]) == True:
            cm.msg("p" , pr_rq.msg_re["pass_approver_save"])
            cm.scroll()
            to_element = driver.find_element_by_xpath( pr_rq.rq_vc["bt_select_cc"])
            cm.scrolling_to_target(to_element)
        
            list_app = driver.find_elements_by_xpath( pr_rq.rq_vc["list_approver1"])
            total_app = cm.total_data(list_app)
        
            if total_app == 0:
                cm.xlsx( pr_rq.re_vc["fail"] , pr_rq.msg_re["fail_approver"])
            else:
                for i in range(1,total_app+1):
                    app_name = driver.find_element_by_xpath(xpath.sl_approver(i)).text
                    if app_name.strip() == user_name.strip():
                        cm.msg("p" , pr_rq.msg_re["pass_approver"])
                        select_approver["result_approver"] = True
                        select_approver["approver_name"] = user_name.strip()
                        break
                        
                if select_approver["result_approver"] == False :
                    cm.xlsx( pr_rq.re_vc["fail"] , pr_rq.msg_re["fail_approver"])
        else:
            cm.xlsx( pr_rq.re_vc["fail"] , pr_rq.msg_re["fail_approver_save"])

    def check_approver_reason(type_request,approver,info_after,info_before):
        
        content_approver = False
        result = xpath.par_approver_result()
        if type_request == "all":

            # Check resaon #
            cm_fu.check_reason(info_before,info_after,result,type_request)

            # Check approver #
            if approver["approval_exception"] == True:
                info_before["approver"] = pr_rq.msg_re["approval_exception"]

                if cm.is_Displayed(pr_rq.rq_vc["approval_exception"]) == False:
                    cm.msg("p" ,type_vc.check_type(type_request , pr_rq.msg_re["pass_detail_approver"]))
                    result["rs_approver"] = True
                    result["approver"] = pr_rq.msg_re["approval_exception"]
                    
                else:
                    cm.xlsx( pr_rq.re_detail["fail"] ,type_vc.check_type(type_request , pr_rq.msg_re["fail_detail_approver"]))

            else:
                i = j = 1
                info_before["approver"] = approver["approver_name"]
                if cm.is_Displayed( pr_rq.rq_vc["approval_exception"]) == True:
                    time.sleep(2)
                    list_approver = driver.find_elements_by_xpath( pr_rq.rq_vc["content_vc_approver"])
                    total_approver = cm.total_data(list_approver)

                    if approver["result_approver"] == True :
                        while i <= total_approver :
                            approver_name = driver.find_element_by_xpath(xpath.approver_name(i)).text
                            if approver_name.strip() in approver["approver_name"]:
                                cm.msg("p" ,type_vc.check_type(type_request , pr_rq.msg_re["pass_view_approver"]))
                                info_after["approver"] = approver["approver_name"]
                                result["rs_approver"] = True
                                content_approver = True
                                break
                            i = i + 1
                        
                        if content_approver == False :
                            cm.xlsx( pr_rq.re_detail["fail"] , type_vc.check_type(type_request , pr_rq.msg_re["fail_view_approver"]))
                    else:
                        while j <= total_approver :
                            approver_name = driver.find_element_by_xpath(xpath.approver_name(j)).text
                            if approver_name == approver["approver_name"][0]:
                                cm.msg("p" , type_vc.check_type(type_request , pr_rq.msg_re["pass_detail_line"]))
                                info_after["approver"] = approver["approver_name"]
                                result["rs_approver"]  = True
                                content_approver       = True
                                break
                            j = j + 1
                        if content_approver == False :
                            cm.xlsx( pr_rq.re_detail["fail"] ,type_vc.check_type(type_request , pr_rq.msg_re["fail_detail_line"]))
                else:
                    cm.xlsx( pr_rq.re_detail["fail"] ,type_vc.check_type(type_request , pr_rq.msg_re["fail_view_approver"]))
            result["info_before"] = info_before
            result["info_after"]  = info_after
            return result

    def add_cc():
        if cm.is_Displayed( pr_rq.rq_vc["bt_add_cc"]) == True:
            cm.msg("p" , pr_rq.msg_re["pass_cc_select"])
            driver.find_element_by_xpath( pr_rq.rq_vc["dele_all_cc"]).click()
            return True
        else :
            cm.xlsx( pr_rq.re_vc["fail"] , pr_rq.msg_re["fail_cc_select"])
            return False
    
    def selected_cc(selected_cc):
        if selected_cc != False :
            cm.msg("p" , pr_rq.msg_re["pass_cc_click_user"])
            time.sleep(1)
            driver.find_element_by_xpath( pr_rq.rq_vc["bt_add_cc"]).click()
            cm.msg("p" , pr_rq.msg_re["pass_cc_add"])

            driver.find_element_by_xpath( pr_rq.rq_vc["bt_save"]).click()
            cm.msg("p" , pr_rq.msg_re["pass_cc_save"])

            time.sleep(1) 
            return True
        else :
            cm.xlsx( pr_rq.re_vc["fail"] , pr_rq.msg_re["fail_cc_click_user"])
            return False

    def check_saved_cc(selected_cc):
        if cm.is_Displayed( pr_rq.rq_vc["bt_select_cc"]) == True :
            cm.msg("p" , pr_rq.msg_re["pass_cc_save_selected"])

            list_cc = driver.find_elements_by_xpath( pr_rq.rq_vc["list_cc"])
            total_cc = cm.total_data(list_cc)

            if total_cc == 0:
                cm.xlsx( pr_rq.re_vc["fail"] , pr_rq.msg_re["fail_cc"])
            else:
                i = 1
                result_cc = False
                while i <= total_cc :
                    cc_name = driver.find_element_by_xpath(xpath.cc_name(i)).text
                    if cc_name.strip() == selected_cc.strip():
                        cm.msg("p" , pr_rq.msg_re["pass_cc"])
                        result_cc = True
                        break
                    i = i+1
                if result_cc == False :
                    cm.xlsx( pr_rq.re_vc["fail"] , pr_rq.msg_re["fail_cc"])
                cm_fu.info_cc(cc_name ,selected_cc)
        else:
            cm.xlsx( pr_rq.re_vc["fail"] , pr_rq.msg_re["fail_cc_save_selected"])
    
    def view_vacation_date(info_before,vc_rq,type_request,info_after):
        result_vc_date = False
        try:
            info_before["vc_date"] = vc_rq["vc_date"]
            vc_date = driver.find_element_by_xpath( pr_rq.rq_vc["content_vc_date"]).text

            if type_request =="all":
                vc_date = xpath.vacation_date("d",vc_date)

            elif type_request == "vc_con":
                vc_date = vc_date
                vc_rq["vc_date"] = vc_rq["vc_date"][None:21]

            elif type_request == "hour":
                vc_date = xpath.vacation_date("d",vc_date)

            else:
                vc_date = xpath.vacation_date("o",vc_date) 

            info_after["vc_date"] = vc_date
            if vc_date == vc_rq["vc_date"] :
                cm.msg("p" , type_vc.check_type(type_request , pr_rq.msg_re["pass_detail_date"]))
                result_vc_date = True
            else:
                cm.xlsx( pr_rq.re_detail["fail"] , type_vc.check_type(type_request , pr_rq.msg_re["fail_detail_date"]))
        except:
            pass
        return result_vc_date

    def view_detail_number_of_days_used(info_before,use_hour_unit,type_request,info_after):
        result_number_use = False
        try :
            use = driver.find_element_by_xpath( pr_rq.rq_vc["content_vc_use"]).text
            days_use = cm_fu.view_detail_used(type_request,use_hour_unit)
            info_after["used"]  = use
            info_before["used"] = days_use
            if use == days_use :
                cm.msg("p" , type_vc.check_type(type_request , pr_rq.msg_re["pass_detail_used"]))
                result_number_use = True
            else:
                cm.xlsx( pr_rq.re_detail["fail"] ,type_vc.check_type(type_request , pr_rq.msg_re["fail_detail_used"]))
        except:
            pass

        return result_number_use

    def view_detail_request_date(info_before,vc_rq,type_request,info_after):
        result_re_date = False
        try :
            info_before["request_date"] = vc_rq["request_date"]
            request_date = driver.find_element_by_xpath( pr_rq.rq_vc["content_request_date"]).text
            info_after["request_date"] = request_date
            if request_date == vc_rq["request_date"] :
                cm.msg("p" , type_vc.check_type(type_request , pr_rq.msg_re["pass_request_date"]))
                result_re_date = True
            else:
                cm.xlsx( pr_rq.re_detail["fail"] ,type_vc.check_type(type_request , pr_rq.msg_re["fail_request_date"]))
        except:
            pass
        return result_re_date
    
    def vacation_date_is_used(list_vc):
        for vc_date in list_vc :
            if vc_date["status"] == "Request" or vc_date["status"] == "Approved" :
                if len(vc_date["vc_date"]) > 0 :
                    date = vc_date["vc_date"][0]
                else :
                    date = vc_date["vc_date"]
                request_date =  datetime.datetime.strptime(date.replace("-", "/"),"%Y/%m/%d")
                driver.find_element_by_link_text("Request Vacation").click()
                return request_date

    
    def view_detail_approver_and_reason(type_request,approver,info_after):
        rs_resaon = rs_approver = False
        try :
            approver_reason = cm_fu.check_approver_reason(type_request,approver,info_after,info_after)
            rs_resaon   = approver_reason["rs_resaon"]
            rs_approver = approver_reason["rs_approver"]
            info_before = approver_reason["info_before"]
            info_after  = approver_reason["info_after"]  
        except:
            pass
        return rs_resaon , rs_approver , info_before , info_after

class cm_rq():
    def hours_set_from_time_card(type_request):
        # Specific working hours from time card  #  
        hour_use = driver.find_element_by_xpath(pr_rq.rq_vc["hour_use"]).text
        hour_use = xpath.tc_hour_use(hour_use,"d")

        if type_request == "all" :
            if len(hour_use) == 0 :
                return 8
            else :
                hour_use = xpath.tc_hour(hour_use)
                if hour_use == 1 :
                    return 8
                else :
                    return hour_use
        elif type_request == "hour":
            hour_use = driver.find_element_by_xpath(pr_rq.rq_vc["hour_use_h"]).text
            hour_use =  xpath.tc_hour_use(hour_use,"h")
            return xpath.tc_hour(hour_use) 
        else:
            return 4   
    
    def select_date_to_request_leave_for_vacation_consecutive():
        i         = 1 
        date_used = [] 
        list_date = []
        
        # Go to my vacation to take used date # 
        cm.msg("n","Select Date")
        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.LINK_TEXT,"My Vacation Status"))).click()
        time.sleep(3)

        rows = driver.find_elements_by_xpath(pr_rq.rq_vc["list_request"])
        total_request = cm.total_data(rows)
        while i <= total_request:
            if i == 1:
                if cm.is_Displayed(pr_rq.rq_vc["check_list_re"]) == True :
                    break
                else:
                    date = driver.find_element_by_xpath(xpath.sd_date(i)).text
                    cm_fu.split_date_from_continuous_date(date,date_used)
            else:
                date = driver.find_element_by_xpath(xpath.sd_date(i)).text
                cm_fu.split_date_from_continuous_date(date,date_used)
            i=i+1

        
        date       = datetime.date.today() 
        start_date = cm_fu.choose_start_date(date_used,date)
        end_date   = cm_fu.choose_end_date(start_date,date_used)
        if end_date == False:
            while end_date == False :
                end_date = cm_fu.choose_end_date(start_date,date_used)
                if end_date != False :
                    list_date.append(start_date)
                    list_date.append(end_date)
                    break
                start_date = cm_fu.next_date(start_date)
                start_date = cm_fu.choose_start_date(date_used,start_date)
        else:
            list_date.append(start_date)
            list_date.append(end_date)

        # Select date from find for request vacation # 
        driver.find_element_by_link_text("Request Vacation").click()
        for i in list_date:
            request_date  = i
            current_month = driver.find_element_by_xpath(pr_rq.rq_vc["current_month"]).text[5:None] 
            
            if int(request_date.month) < 10 :
                request_month = "0" + str(request_date.month)
            else :
                request_month = request_date.month
            
            if str(current_month) == str(request_month) :
                result_click = cm_fu.click_date(request_date)
                if result_click == True:
                    cm.msg("p" , pr_rq.msg_re["pass_select_date"])
                else:
                    cm.msg("p" , pr_rq.msg_re["fail_select_date"])
            else:
                driver.find_element_by_xpath(pr_rq.rq_vc["icon_next_month"]).click()
                result_click = cm_fu.click_date(request_date)
                if result_click == True:
                    cm.msg("p" , pr_rq.msg_re["pass_select_date"])
                else:
                    cm.msg("p" , pr_rq.msg_re["fail_select_date"])
            time.sleep(2)

        return list_date

    def select_date_to_request_leave():
        i = 1 
        date_used = [] 
        
        # Go to my vacation to take used days # 
        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.LINK_TEXT,"My Vacation Status"))).click()
        time.sleep(3)
        rows          = driver.find_elements_by_xpath(pr_rq.rq_vc["list_request"])
        total_request = cm.total_data(rows)
        while i <= total_request:
            if i == 1:
                if cm.is_Displayed(pr_rq.rq_vc["check_list_re"]) == True :
                    break
                else:
                    date = driver.find_element_by_xpath(xpath.sd_date(i)).text
                    cm_fu.split_date_from_continuous_date(date,date_used)
            else:
                date = driver.find_element_by_xpath(xpath.sd_date(i)).text
                cm_fu.split_date_from_continuous_date(date,date_used)
            i = i+1
        
    
        # Find unused days , not saturday , not sunday , not holiday to use for request vacation # 
        request_date = datetime.date.today() 
        if request_date.weekday() == 5 :
            request_date = request_date + relativedelta(day=request_date.day +2)
        if  request_date.weekday() == 6 :
            request_date = request_date + relativedelta(day=request_date.day +1)
        if str(request_date) in date_used  :
            request_date = request_date + relativedelta(day=request_date.day +1)
            if request_date.weekday() == 5 :
                request_date = request_date + relativedelta(day=request_date.day +2)
            if  request_date.weekday() == 6 :
                request_date = request_date + relativedelta(day=request_date.day +1)
            
            while str(request_date) in date_used  :
                request_date = cm_fu.next_date(request_date)
                if request_date.weekday() == 5 :
                    request_date = request_date + relativedelta(day=request_date.day +2)
                if  request_date.weekday() == 6 :
                    request_date= request_date + relativedelta(day=request_date.day +1)
                
        # Select date from find for request vacation # 
        driver.find_element_by_link_text("Request Vacation").click()
        cm.popup_time_card()
        current_month = driver.find_element_by_xpath(pr_rq.rq_vc["current_month"]).text[5:None] 
        if int(request_date.month) < 10 :
            request_month = "0" + str(request_date.month)
        else :
            request_month = request_date.month

        if str(current_month) == str(request_month) :
            result_click = cm_fu.click_date(request_date)
            if result_click == True:
                cm.msg("p" , pr_rq.msg_re["pass_select_date"])
                return request_date
            else:
                cm.msg("p" , pr_rq.msg_re["fail_select_date"])
                return False
        else:
            driver.find_element_by_xpath(pr_rq.rq_vc["icon_next_month"]).click()
            result_click = cm_fu.click_date(request_date)
            if result_click == True:
                cm.msg("p" , pr_rq.msg_re["pass_select_date"])
                return request_date
            else:
                cm.msg("p" , pr_rq.msg_re["fail_select_date"])
                return False
    
    def get_vacation_date_and_status():
        i = 1 
        list_vc   = [] 
        
        # Go to my vacation to take used days # 
        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.LINK_TEXT,"My Vacation Status"))).click()
        time.sleep(3)
        rows          = driver.find_elements_by_xpath(pr_rq.rq_vc["list_request"])
        total_request = cm.total_data(rows)
        while i <= total_request:
            if i == 1:
                if cm.is_Displayed(pr_rq.rq_vc["check_list_re"]) == True :
                    break
                else:
                    date    = driver.find_element_by_xpath(xpath.sd_date(i)).text
                    vc_date = cm_fu.get_vacation_date(date)
                    vc_stat = driver.find_element_by_xpath(xpath.sd_status(i)).text
                    list_vc.append(xpath.par_date_and_sta(vc_date,vc_stat))

            else:
                date    = driver.find_element_by_xpath(xpath.sd_date(i)).text
                vc_date = cm_fu.get_vacation_date(date)
                vc_stat = driver.find_element_by_xpath(xpath.sd_status(i)).text
                list_vc.append(xpath.par_date_and_sta(vc_date,vc_stat))
            i = i+1

        return list_vc

           
    def available_vacation():
    
        # Get data of each vacation at available vacation table #  
        i = 1 
        total_vacation = 0 
        all_vacation   = []

        driver.find_element_by_link_text("Request Vacation").click()
        tbody = driver.find_element_by_xpath(data["available_vacation"]["tbody"])
        rows  = tbody.find_elements_by_tag_name("tr")
        total_vacation = cm.total_data(rows)

        while i <= total_vacation:
            vacation = xpath.par_vacation()
            total_days = driver.find_element_by_xpath(xpath.vc_available(i,"to")).text
            if total_days != "-" :
                vacation_name      = driver.find_element_by_xpath(xpath.vc_available(i,"na")).text
                vacation["total"]  = total_days
                vacation["used"]   = driver.find_element_by_xpath(xpath.vc_available(i,"us")).text 
                vacation["remain"] = driver.find_element_by_xpath(xpath.vc_available(i,"re")).text 
                vacation["expiration"] = driver.find_element_by_xpath(xpath.vc_available(i,"ex")).text 
                vacation["start"]      = xpath.vc_start(vacation_name)
                vacation["vacation_name"] = xpath.vc_change(vacation, vacation_name)
                all_vacation.append(vacation)
            else :
                vacation["vacation_name"] = driver.find_element_by_xpath(xpath.vc_available(i,"na")).text
                vacation["total"] =  vacation["remain"] = vacation["start"] = vacation["expiration"] = "-"
                vacation["used"]  = driver.find_element_by_xpath(xpath.vc_available(i,"us")).text 
                all_vacation.append(vacation)
            i = i+1 
        return all_vacation

    def total_vacation():
        
        # Total vacation availabel of user #  
        i = 1 
        total_vacation_can_use = 0
        driver.find_element_by_link_text("Request Vacation").click()
        time.sleep(3)

        tbody = driver.find_element_by_xpath(data["available_vacation"]["tbody"])
        rows  = tbody.find_elements_by_tag_name("tr")
        total_vacation = cm.total_data(rows)
        while i <= total_vacation:
            remain = driver.find_element_by_xpath(xpath.total_remain(i)).text 
            if str(remain.strip()) != "0" :
                total_vacation_can_use += 1
            i = i+1
    
        return total_vacation_can_use

    def check_number_of_days_off(before,after,hour_use,vc_name,oneday,use_hour_unit,type_request):

        cm.msg("p" , "  Available Vacation")
        infor_before = infor_after = " " 
        total = used = remain = True 
        numberex = pr_rq.param_excel_re["number"]

        # Check number of days before request #
        for vacation in before:
            if vacation["vacation_name"] == vc_name :
                vc_bf_use    = vacation
                infor_before = cm_fu.infor(vc_bf_use,"Info Vacation before request",type_request)
                total_before = param["total_before"]

                if vc_bf_use["total"] != "-" :
                    days = cm_fu.change_hour_to_day(vc_bf_use["used"],vc_bf_use["remain"],oneday,"plus",hour_use,use_hour_unit,type_request)
                    cm_fu.result(vc_bf_use["total"],days,total_before,numberex,type_request)
                else :
                    cm_fu.result(vc_bf_use["total"],"-",total_before,numberex,type_request)
                break

        # Check number of days after request # 
        for vacation in after: 
            if vacation["vacation_name"] == vc_name :
                vc_af_use   = vacation
                infor_after = cm_fu.infor(vc_af_use,"Info Vacation after request ",type_request)
                total_list  = param["total_list"]
                used_list   = param["used_list"]
                remain_list = param["remain_list"]

                # vacation type is regular / grant #
                if vc_bf_use["total"] != "-" : 
                    total = cm_fu.result_number(vc_af_use["total"],vc_bf_use["total"],total_list,numberex,total,type_request)
                
                    used_before_plus_used = cm_fu.change_hour_to_day(vc_bf_use["used"],None,oneday,"plus",hour_use,use_hour_unit,type_request)
                    used = cm_fu.result_number(vc_af_use["used"],used_before_plus_used,used_list,numberex,used,type_request)

                    remain_before_minus_used = cm_fu.change_hour_to_day(vc_bf_use["remain"],None,oneday,"minus",hour_use,use_hour_unit,type_request)
                    remain = cm_fu.result_number(vc_af_use["remain"],remain_before_minus_used,remain_list,numberex,remain,type_request)
                    
                    all_list = param["all_list"]
                    cm_fu.result_all(total,used,remain,all_list,numberex,type_request)

                # vacation type is Other #
                else :
                    total = cm_fu.result_number(vc_af_use["total"],"-",total_list,numberex,total,type_request)
                    used_before_plus_used =  cm_fu.change_hour_to_day(vc_bf_use["used"],None,oneday,"plus",hour_use,use_hour_unit,type_request)
                    used     = cm_fu.result_number(vc_af_use["used"],used_before_plus_used,used_list,numberex,used,type_request)
                    remain   = cm_fu.result_number(vc_af_use["remain"],"-",remain_list,numberex,remain,type_request)
                    all_list = pr_rq.param_excel_re["all_list"]
                    cm_fu.result_all(total,used,remain,all_list,numberex,type_request)

                break
        
        cm.msg("t" ,infor_before)
        cm.msg("t" ,infor_after)
        
    def check_number_of_days_cancel(before,after,hour_use,vc_name,oneday,use_hour_unit,type_request):
        cm.msg("p","  Available Vacation")
        total = used = remain = True 
        numberex =  pr_rq.param_excel_re["number"]

        # Check number of days before request #
        for vacation in before:
            if vacation["vacation_name"] == vc_name :
                vc_bf_use     = vacation
                infor_before  = cm_fu.infor(vc_bf_use,"Info Vacation before cancel",type_request)
                cancel_before = param["cancel_before"]

                if vc_bf_use["total"] != "-" :
                    days = cm_fu.change_hour_to_day(vc_bf_use["used"],vc_bf_use["remain"],oneday,"plus",hour_use,use_hour_unit,type_request)
                    cm_fu.result(vc_bf_use["total"],days,cancel_before,numberex,type_request)
                else :
                    cm_fu.result(vc_bf_use["total"],"-",cancel_before,numberex,type_request)
                break

        # Check number of days after request #
        for vacation in after: 
            if vacation["vacation_name"] == vc_name :
                vc_af_use    = vacation
                infor_after  = cm_fu.infor(vc_af_use,"Info Vacation after cancel",type_request)
                total_cancel = param["total_cancel"]
                used_cancel  = param["used_cancel"]
                remain_cancel= param["remain_cancel"]
                if vc_bf_use["total"] != "-" : 
                    
                    total = cm_fu.result_number(vc_af_use["total"],vc_bf_use["total"],total_cancel,numberex,total,type_request)
                
                    used_before_plus_used = cm_fu.change_hour_to_day(vc_bf_use["used"],None,oneday,"minus",hour_use,use_hour_unit,type_request)
                    used = cm_fu.result_number(vc_af_use["used"],used_before_plus_used,used_cancel,numberex,used,type_request)

                    remain_before_minus_used = cm_fu.change_hour_to_day(vc_bf_use["remain"],None,oneday,"plus",hour_use,use_hour_unit,type_request)
                    remain = cm_fu.result_number(vc_af_use["remain"],remain_before_minus_used,remain_cancel,numberex,remain,type_request)
                    
                    all_cancel = param["all_cancel"]
                    cm_fu.result_all(total,used,remain,all_cancel,numberex,type_request)

                else :
                    total = cm_fu.result_number(vc_af_use["total"],"-",total_cancel,numberex,total,type_request)

                    used_before_plus_used = cm_fu.change_hour_to_day(vc_bf_use["used"],None,oneday,"minus",hour_use,use_hour_unit,type_request)
                    used = cm_fu.result_number(vc_af_use["used"],used_before_plus_used,used_cancel,numberex,used,type_request)

                    remain = cm_fu.result_number(vc_af_use["remain"],"-",remain_cancel,numberex,remain,type_request)

                    all_cancel = param["all_cancel"]
                    cm_fu.result_all(total,used,remain,all_cancel,numberex,type_request)
                break

        cm.msg("t" ,infor_before)
        cm.msg("t" ,infor_after)

    def select_approver():
    
        cm.scroll()
         
        bt_select_approver = cm.is_Displayed( pr_rq.rq_vc["bt_select_approver"])
        select_approver = xpath.par_select_approver()
        if bt_select_approver == True:
            # Select approver from approver list #
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,  pr_rq.rq_vc["bt_select_approver"]))).click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,  pr_rq.rq_vc["delete_all"]))).click()
            time.sleep(3)

            check_no_approver = driver.find_elements_by_xpath( pr_rq.rq_vc["text_list_ap"])
            if cm_fu.approver_list(check_no_approver) == True :
                selected_approver = cm_fu.select_approver()
                if selected_approver["is_selected"] == True :
                    if cm_fu.add_approver() == True :
                        cm_fu.save_approver(selected_approver["user_name"],select_approver)
            return select_approver

        else:
            # Use approver line #
            if cm.is_Displayed( pr_rq.rq_vc["bt_quick_approver"]) == True: 
                cm.msg("p" , pr_rq.msg_re["pass_approver_line"])
                select_approver["result_approver"] = True
                select_approver["approval_line"]   = True
                list_approver = []
                list_app      = driver.find_elements_by_xpath( pr_rq.rq_vc["list_approver1"])
                total_app     = cm.total_data(list_app)
                select_approver["approver_name"] = cm_fu.approver_name(total_app,list_approver)
                return select_approver 

            else:
                # Approver is approval exception #
                cm.msg("p" , pr_rq.msg_re["pass_approver_exception"])
                select_approver["result_approver"]    = True
                select_approver["approval_exception"] = True
                return select_approver

    def function_search():
        try:
            user_name = "TS2"
            result_select_user = {
                "search"    :"False",
                "select_ogr":"False"
                }
            driver.find_element_by_xpath( pr_rq.rq_vc["bt_select_cc"]).click()
            if cm.is_Displayed( pr_rq.rq_vc["bt_add_cc"]) == True:
                driver.find_element_by_xpath( pr_rq.rq_vc["dele_all_cc"]).click()

                # Search user # 
                list_departmaent = driver.find_elements_by_xpath( pr_rq.rq_vc["org_search"])
                before_search    = len(list_departmaent)
                firt_department  = driver.find_element_by_xpath( pr_rq.rq_vc["firt_depart"]).text
                ip_search_user   = driver.find_element_by_xpath( pr_rq.rq_vc["search"])
                driver.implicitly_wait(5)
                ip_search_user.click()
                ip_search_user.send_keys(user_name)
                ip_search_user.send_keys(Keys.RETURN)
                
                if ip_search_user.get_attribute('value') == user_name :
                    cm.msg("p" , pr_rq.msg_re["pass_approver_enter_user"])
                    
                    time.sleep(1)
                    cc = driver.find_element_by_xpath( pr_rq.rq_vc["cc_namea"]).text
                    if cc == "No data." :
                        cm.xlsx( pr_rq.re_search["pass"] , pr_rq.msg_re["pass_cc_search_user"])

                    else:
                        time.sleep(2)
                        list_depart = driver.find_elements_by_xpath( pr_rq.rq_vc["org_search"])
                        after_search = len(list_depart)
                        if before_search != after_search :
                            cm.xlsx( pr_rq.re_search["pass"] , pr_rq.msg_re["pass_cc_search_user"])
                            result_select_user["search"]=True
                        else:
                            firt_department1 = driver.find_element_by_xpath(pr_rq.rq_vc["firt_depart"]).text
                            if firt_department == firt_department1:
                                cm.xlsx( pr_rq.re_search["fail"] , pr_rq.msg_re["fail_cc_search_user"])
                            else:
                                cm.xlsx( pr_rq.re_search["pass"] , pr_rq.msg_re["pass_cc_search_user"])
                                result_select_user["search"] = True      
                else:
                    cm.xlsx( pr_rq.re_search["fail"] , pr_rq.msg_re["fail_approver_enter_user"])

                # Select user from Org # 
                ip_search_user.clear()
                driver.find_element_by_xpath( pr_rq.rq_vc["bt_save"]).click()
                driver.find_element_by_xpath( pr_rq.rq_vc["bt_select_cc"]).click()
                selected_cc = cm_fu.select_user_from_depart() 
                if selected_cc != False :
                    cm.xlsx( pr_rq.re_org["pass"] , pr_rq.msg_re["pass_org"])
                    result_select_user["select_ogr"] = True      
                else:
                    cm.xlsx( pr_rq.re_org["fail"] , pr_rq.msg_re["fail_org"])
                driver.find_element_by_xpath( pr_rq.rq_vc["bt_save"]).click()
            
        except:
            driver.find_element_by_link_text("My Vacation Status").click()

    def select_cc_enter_reason():
        # SELECT CC #
        time.sleep(3)
        cm.scroll()
        cm.msg("n" , "Select CC")
        driver.find_element_by_xpath( pr_rq.rq_vc["bt_select_cc"]).click()
        if cm_fu.add_cc() == True :
            selected_cc = cm_fu.select_user_from_depart().replace("(", "").replace(")", "")
            if cm_fu.selected_cc(selected_cc) == True :
                cm_fu.check_saved_cc(selected_cc)

        # ADD REASON #
        cm.scroll()
        cm.msg("n" ,"Enter Reason")
        if cm.is_Displayed( pr_rq.rq_vc["reason"]) == True:
            reason=driver.find_element_by_xpath( pr_rq.rq_vc["reason"])
            reason.click()
            reason.send_keys( pr_rq.rq_vc["reason_text"])
            if reason.get_attribute('value') ==  pr_rq.rq_vc["reason_text"]:
                cm.msg("p" , pr_rq.msg_re["pass_reason"])
            else:
                cm.xlsx( pr_rq.re_vc["fail"] , pr_rq.msg_re["fail_reason"])
        else:
            cm.msg("p" , pr_rq.msg_re["pass_reason_no"])
        
    def check_use_hour_unit_half_day(total_vc):
        # Choose vacation name to request  #
        use_hour_unit      = False 
        all_vacation       = [] 
        available_vacation = {"available_vacation":""}
        
        if total_vc  == 0 :
            available_vacation["available_vacation"] = 0
            all_vacation.append(available_vacation)

        else:
            i = 1
            available_vacation["available_vacation"] = total_vc
            all_vacation.append(available_vacation)
            while i <= total_vc:
                time.sleep(1)
                usage_settings = xpath.par_usage_settings()
                driver.find_element_by_css_selector( pr_rq.rq_vc["select_vacation"]).click()
                driver.find_element_by_xpath("//body/div[4]/div/div/div[" + str(i) +"]").click()
                vacation = cm_fu.vacation_use_for_request()
                usage_settings["vacation_name"]   = vacation["vacation_name"]
                usage_settings["number_of_days"]  = vacation["number_of_days"]
                usage_settings["number_of_hours"] = vacation["number_of_hours"]

                if cm.is_Displayed( pr_rq.rq_vc["hour_unit"]) == True :
                    use_hour_unit = True
                    usage_settings["use_hour_unit"] = True
                else:
                    usage_settings["use_hour_unit"] = False

                if cm.is_Displayed( pr_rq.rq_vc["radi_am"]) == True:
                    usage_settings["use_half_day"] = True
                    if use_hour_unit == True:
                        usage_settings["hour_use"] = cm_fu.hour_used(use_hour_unit,"am")
                else:
                    usage_settings["use_half_day"] = False

                all_vacation.append(usage_settings)
                i = i+1
        return all_vacation    
                
    def select_vacation_use_hour_unit_half_day(total_vc,list_vc_use_half,hour_use,type_vc):
        i = 1
        while i <= total_vc:
            time.sleep(1)
            driver.find_element_by_css_selector( pr_rq.rq_vc["select_vacation"]).click()
            driver.find_element_by_xpath("//body/div[4]/div/div/div[" + str(i) +"]").click()
            vacation_name = driver.find_element_by_xpath( pr_rq.rq_vc["vacation_name"]).text
            vacation_name = xpath.us_change(vacation_name)
            for vacation in list_vc_use_half:
                if vacation["vacation_name"] == vacation_name :
                    if type_vc =="hour":
                        if  float(vacation["number_of_hours"]) >= 1       or \
                            float(vacation["number_of_days"] ) >= hour_use:
                            return vacation_name
                    else:
                        if  float(vacation["number_of_hours"]) >= hour_use or \
                            float(vacation["number_of_days"] ) >= hour_use:
                            return vacation_name
            i = i+1      
                            
        return False

    def check_result_request():
        try :
            notification = driver.execute_script('return document.getElementById("noty_layout__topRight").innerText')
            content_notification = notification.split("\n")
            if content_notification[0] =="success":
                return "pass"
            else:
                return "noti_error" + content_notification[1]
        except:
            time.sleep(1)
            if cm.is_Displayed(data["my_vt"]["vc_history"]) == True:
                return "pass"
            else:
                return "fail"
      
    def created_request_and_view_detail(info_vc,type_request,use_hour_unit,approver):
        i      = 1
        result = False
        info_before = xpath.par_info()
        info_after  = xpath.par_info()

        time.sleep(3)
        try:
            driver.find_element_by_xpath( pr_rq.rq_vc["bt_refresh"]).click()
            list_request = driver.find_elements_by_xpath( pr_rq.rq_vc["list_request"])
            total_request = cm.total_data(list_request)
            
            if total_request >= 1 :
                while i <= total_request:
                    vc_rq = xpath.par_vc_rq()
                    if approver["approval_exception"] == True:
                        vc_rq["status"]   = "Approved"
                        info_vc["status"] = "Approved"

                    
                    vc_rq = cm_fu.vacation_request(vc_rq,i)
                    if  info_vc["vc_name"] == vc_rq["vc_name"]           and \
                        info_vc["vc_date"] == vc_rq["vc_date"]           and \
                        info_vc["request_date"] == vc_rq["request_date"] and \
                        info_vc["status"]  == vc_rq["status"] :
                        cm.xlsx( pr_rq.re_vc["pass"] , type_vc.check_type(type_request, pr_rq.msg_re["pass_request_displayed"]))
                        cm.msg("n" ,"View Detail")
                        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH , xpath.view_detail(i)))).click()
                        time.sleep(2)
                        

                        # View vacation date #
                        result_vc_date = cm_fu.view_vacation_date(info_before,vc_rq,type_request,info_after)
                        
                        # View detail Number of days used #
                        result_number_use = cm_fu.view_detail_number_of_days_used(info_before,use_hour_unit,type_request,info_after)
                        
                        # View detail request date #
                        result_re_date = cm_fu.view_detail_request_date(info_before,vc_rq,type_request,info_after)
                        
                        # View detail approver and reason #
                        rs_resaon , rs_approver , info_before , info_after = cm_fu.view_detail_approver_and_reason(type_request,approver,info_after)
                        if type_request == "all":
                            if  result_vc_date    == True and \
                                result_number_use == True and \
                                result_re_date    == True and \
                                rs_resaon         == True and \
                                rs_approver       == True     :
                                cm.xlsx( pr_rq.re_detail["pass"] , type_vc.check_type(type_request , pr_rq.msg_re["pass_detail"]))
                            else:
                                cm.xlsx( pr_rq.re_detail["fail"] , type_vc.check_type(type_request , pr_rq.msg_re["fail_detail"]))
                        else:
                            if  result_vc_date     == True and \
                                result_number_use  == True and \
                                result_re_date     == True     :
                                cm.xlsx( pr_rq.re_detail["pass"] , type_vc.check_type(type_request , pr_rq.msg_re["pass_detail"]))
                            else:
                                cm.xlsx( pr_rq.re_detail["fail"] , type_vc.check_type(type_request , pr_rq.msg_re["fail_detail"]))


                        cm_fu.infor_detail("Information entered",info_before)
                        cm_fu.infor_detail("Information saved  ",info_after )
                        result = True
                        break

                    i = i+1

                if result == False:
                    cm.xlsx( pr_rq.re_vc["fail"] , type_vc.check_type(type_request, pr_rq.msg_re["fail_request_displayed"]))
            else: 
                cm.xlsx( pr_rq.re_vc["fail"] , type_vc.check_type(type_request, pr_rq.msg_re["fail_request_displayed"]))
           
        except:
            driver.find_element_by_link_text("My Vacation Status").click()

    def time_comparison(request_date,today):
    
        if request_date.rfind("~") > 0 :
            request_date = request_date[int(request_date.rfind("~"))+1: None]

        request_date = datetime.datetime.strptime(request_date.replace("-", "/").replace("2022", "22"), "%y/%m/%d")
        today        = datetime.datetime.strptime(today.replace("-", "/").replace("2022", "22"), "%y/%m/%d")
        if request_date < today :
            return False  
        else:
            return True

    def select_hour_use_hour_unit():
        
        cm.msg("n" ,"Select Hour")
        driver.find_element_by_xpath( pr_rq.rq_vc["hour_unit"]).click()
        driver.find_element_by_xpath( pr_rq.rq_vc["hour_start"]).click()
        
        start_options = driver.find_elements_by_xpath( pr_rq.rq_vc["start_option"])
        if len(start_options) > 1:
            driver.find_element_by_xpath( pr_rq.rq_vc["sl_hour_start"]).click()
        else:
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_hour_start"])
            return False

        driver.find_element_by_xpath( pr_rq.rq_vc["hour_end"]).click()
        end_options = driver.find_elements_by_xpath( pr_rq.rq_vc["end_option"])
        if len(end_options) > 1:
            driver.find_element_by_xpath( pr_rq.rq_vc["sl_hour_end"]).click()
        else:
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_hour_end"])
            return False
        
        hour_selected = driver.find_element_by_xpath( pr_rq.rq_vc["selected_date"]).text
        hour_selected = hour_selected[int(hour_selected.rfind("(")): int(hour_selected.rfind(")"))+1]
        return hour_selected
        
    def info_request_list(i):
    
        infor_request = xpath.par_infor_request()
        infor_request["no"] = str(i)
        infor_request["vacation_name"] = driver.find_element_by_xpath(xpath.li_request(i,"na")).text
        infor_request["vacation_date"] = driver.find_element_by_xpath(xpath.li_request(i,"da")).text
        infor_request["use"]           = driver.find_element_by_xpath(xpath.li_request(i,"us")).text
        infor_request["request_date"]  = driver.find_element_by_xpath(xpath.li_request(i,"rd")).text
        infor_request["status"]        = driver.find_element_by_xpath(xpath.li_request(i,"st")).text
        infor_request["vacation_time"] = infor_request["vacation_date"][None: int(infor_request["vacation_date"].rfind("\n"))]
        return infor_request

    def count_all_vacation_request():
        # Get all status from list request vacation #
        
        i = 1
        total_request = 0
        time.sleep(3)

        if cm.is_Displayed( pr_rq.rq_vc["check_list_re"]) == False : 
            driver.find_element_by_xpath(data["mn_pro"]["ic_to_end_page"]).click()
            end_page_text = driver.find_element_by_xpath(data["mn_pro"]["page_current"]).text
            end_page = int(end_page_text)
            driver.find_element_by_xpath(data["mn_pro"]["ic_to_first_page"]).click()
            
            while i <= end_page:
                if i == end_page :
                    time.sleep(3)
                    total_re = driver.find_elements_by_xpath(data["mn_pro"]["list_re_vc"])
                    total_request = total_request + cm.total_data(total_re)
                else:
                    total_request = total_request+20
                i = i+1
        return total_request

    def two_requests_are_the_same(request1,request2):
        if  request1["vacation_name"] == request2["vacation_name"] and \
            request1["vacation_date"] == request2["vacation_date"] and \
            request1["use"]           == request2["use"]           and \
            request1["request_date"]  == request2["request_date"]:
            return True
        else: 
            return False
                
    def login(domain):
        
        cm.msg("n" ,"LOGIN")
        driver.get("http://"+domain+"/ngw/app/#/sign")
        driver.implicitly_wait(10)


        driver.find_element_by_id("log-userid").send_keys(data["user"])
        cm.msg("p" , pr_rq.msg_re["pass_login_id"])
    
        driver.switch_to.frame(driver.find_element_by_id("iframeLoginPassword"))
        driver.find_element_by_id("p").send_keys(data["pass"])
        cm.msg("p" , pr_rq.msg_re["pass_login_pw"])
        driver.switch_to.default_content()

        driver.find_element_by_id("btn-log").send_keys(Keys.RETURN)
        cm.msg("p" , pr_rq.msg_re["pass_login_bt"])
        
    def access_menu_vacation(domain):
        
        #window#
        driver.get("http://"+domain+"/ngw/app/#/nhr")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,data["iframe_vc"])))
        driver.switch_to.frame(driver.find_element_by_xpath(data["iframe_vc"]))
        #
        
        if cm.is_Displayed(data["menu_vc"]) == True :
            driver.find_element_by_xpath(data["menu_vc"]).click()
            if cm.is_Displayed(data["sm_my_vc_sta"]) == True :
                result =  True
                cm.msg("n" ,"ACCESS VACATION")
                cm.msg("p" , pr_rq.msg_re["pass_acccess_menu"])
                cm.popup_time_card()
                
            else:
                result = False
                cm.msg("f" , pr_rq.msg_re["fail_acccess_menu"])
        else :
            result = False
            cm.msg("p" , pr_rq.msg_re["pass_no_vacation"])
        return result

    def vacation_displayed_in_time_card(date_request):
        try :
            i = 1
            current_date = datetime.date.today()
            month_request = date_request.month
            current_month = current_date.month
            number_of_clicks = int(month_request - current_month)

            cm.msg("n" ,"Time Card")
            driver.find_element_by_xpath(data["menu_tc"]).click()
            driver.find_element_by_link_text("Timesheets").click()
            
            if cm.is_Displayed(data["tab_calen_tc"]) == True :

                if cm.is_Displayed(data["no_work_tc"]) == False :
                    cm.xlsx( pr_rq.re_timecard["pass"] , pr_rq.msg_re["pass_work_policy"])
                else:
                    time.sleep(2)
                    driver.find_element_by_css_selector(data["date_tc"]).click()
                    if number_of_clicks == 0 :
                        cm_fu.click_date_time_card(date_request)
                    else:
                        while i <= number_of_clicks :
                            driver.find_element_by_css_selector(data["ic_next_tc"]).click()
                            cm_fu.click_date_time_card(date_request)

                    if cm.is_Displayed(data["row_vaca_tc"]) == True :
                        cm.xlsx( pr_rq.re_timecard["pass"] , pr_rq.msg_re["pass_vacation_time_card"])
                    else:
                        cm.xlsx( pr_rq.re_timecard["fail"] , pr_rq.msg_re["fail_vacation_time_card"])

            else:
                cm.xlsx( pr_rq.re_timesheet["fail"] , pr_rq.msg_re["fail_access_time_card"])

            driver.find_element_by_xpath(data["menu_vc"]).click()
        except:
            driver.find_element_by_xpath(data["menu_vc"]).click()

    def time_clockin():
        try :
            time_clock_in =  False
            driver.find_element_by_xpath(data["menu_tc"]).click()
            driver.find_element_by_link_text("Timesheets").click()
            
            if cm.is_Displayed(data["tab_calen_tc"]) == True :
                if cm.is_Displayed(data["no_work_tc"]) == True :
                    clock_in = driver.find_element_by_xpath(data["time_clock_in"]).text
                    if clock_in.rfind("00") > 0 :
                        time_clock_in = clock_in

                    driver.find_element_by_xpath(data["menu_vc"]).click()
                    return time_clock_in           
        except:
            driver.find_element_by_xpath(data["menu_vc"]).click()

    def collect_clock_in_from_time_card(time_clock_in,type_request,request_date):
        if time_clock_in !=  False :
            clock_in = time_clock_in[None: int(time_clock_in.rfind("("))]
            #clock_in_hour = time_clock_in[None: int(time_clock_in.rfind(":"))]
            clock_in_hour = time_clock_in[None: 2]

            if type_request =="hour":
                time_end = int(clock_in_hour) + 2
                vacation_date = str(request_date)+" [ Use: 2H ] "+ "[ " + clock_in + "~ " + str(time_end) +":00 ]"
            else:
                time_end = int(clock_in_hour) + 4 
                vacation_date = str(request_date)+" [ Use: 4H ] "+ "[ " + clock_in + "~ " + str(time_end) +":00 ]"
            return vacation_date

    def information_vacation(title,vacation_request):
        vacation_name = "Vacation Name : "+vacation_request["vc_name"]
        vacation_date = "Vacation Date : "+vacation_request["vc_date"]
        request_date  = "Request Date  : "+vacation_request["request_date"]
        data =  "  +" +title +" [ "+vacation_name +" | " +vacation_date + " | " + request_date +" ] "
        cm.msg("t",data)

    def update_status_for_request(approver,vacation_request):
       
        if approver["approval_exception"] == True :
            vacation_request["status"] = "Approved"
        return vacation_request
    
    def choose_date_to_request(vacation_request,type):
        result_date = True
        if type == "all":
            vc_date = cm_rq.select_date_to_request_leave()
            if vc_date != False:
                vacation_request["vc_date"] = str(vc_date) + "All day Off"
            else:
                result_date = False
        elif type == "am" :
            vacation_request["request_date"] = str(datetime.date.today())
            vc_date = cm_rq.select_date_to_request_leave()
            if vc_date != False:
                vacation_request["vc_date"] = str(vc_date) + "Half Day (AM)"
            else:
                result_date = False
        elif type == "pm" :
            vacation_request["request_date"] = str(datetime.date.today())
            vc_date = cm_rq.select_date_to_request_leave()
            if vc_date != False:
                vacation_request["vc_date"] = str(vc_date) + "Half Day (PM)"
            else:
                result_date = False
        
        elif type == "hour" :
            vacation_request["request_date"] =  str(datetime.date.today())
            vc_date =  cm_rq.select_date_to_request_leave()
            vacation_request["vc_date"] = vc_date
            if vc_date == False:
                result_date = False

        else :
            vc_date = cm_rq.select_date_to_request_leave_for_vacation_consecutive()
            if vc_date != False:
                vacation_request["vc_date"] = str(vc_date[0]) + "~" + str(vc_date[1]) + "All day Off-All day Off"
            else:
                result_date = False
        

        choose_date = xpath.par_choose_date(result_date,vc_date,vacation_request)
        return choose_date
    

    def all_choose_date(all,choose_date):
        all["result_date"]      = choose_date["result_date"] 
        all["vc_date"]          = choose_date["vc_date"]
        all["vacation_request"] = choose_date["vacation_request"]

    def all_choose_vc(all,vacation_request,all_day):
        all["oneday"]                = cm_rq.hours_set_from_time_card("all")
        all["number_before_request"] = cm_rq.available_vacation()
        result_vacation              = all_day.choose_vacation_to_request(vacation_request)
        all["result_select_vc"]      = result_vacation["result_select_vc"]
        all["use_hour_unit"]         = result_vacation["use_hour_unit"]
        all["hour_use"]              = result_vacation["hour_use"]
        all["vc_name"]               = result_vacation["vc_name"]

    def half_choose_vc(vc_half,vacation_request,type,half_day):
        vc_half["number_before_request"] = cm_rq.available_vacation()
        vc_half["oneday"]                = cm_rq.hours_set_from_time_card("all")
        result_vacation                  = half_day.choose_vacation_to_request(vacation_request,type)
        vc_half["result_select_vc"]      = result_vacation["result_select_vc"]
        vc_half["use_hour_unit"]         = result_vacation["use_hour_unit"]
        vc_half["hour_use"]              = result_vacation["hour_use"]
        vc_half["vc_name"]               = result_vacation["vc_name"]

    def init(self,all_day):
        self.result_date = all_day["result_date"]
        self.vc_date     = all_day["vc_date"]
        self.oneday      = all_day["oneday"]
        self.hour_use    = all_day["hour_use"]
        self.vc_name     = all_day["vc_name"]
        self.number_before_request = all_day["number_before_request"]
        self.result_select_vc      = all_day["result_select_vc"]
        self.use_hour_unit         = all_day["use_hour_unit"]
        self.vacation_request      = all_day["vacation_request"] 
        return self

   
class all_day :
    
    def choose_vacation_to_request(vacation_request):
        cm.msg("n" ,"Vacation Name")
        result_select_vc = False
        use_hour_unit    = False
        total_vc         = cm_rq.total_vacation()
        infor_vacation   = cm_rq.check_use_hour_unit_half_day(total_vc)
        
        # Check there are any vacations using hour unit #
        # Change the way to check the number of leave days #
        i=1
        while i <= total_vc:
            if infor_vacation[i]["use_hour_unit"] == True:
                use_hour_unit = True
                break
            i = i + 1

        # Check the remaining days of each vacation are enough to request #
        i = 1
        while i <= total_vc:
            time.sleep(1)
            driver.find_element_by_css_selector( pr_rq.rq_vc["select_vacation"]).click()
            driver.find_element_by_xpath("//body/div[4]/div/div/div[" + str(i) +"]").click()
            vc_use_for_request = cm_fu.vacation_use_for_request()
            hour_use = cm_fu.hour_used(use_hour_unit,"all")
            remain_days = vc_use_for_request["number_of_days"]
            remain_hours = vc_use_for_request["number_of_hours"]
            vc_name = vc_use_for_request["vacation_name"]

            # If vacation is other vacation can use this vacation to request #
            if float(remain_days) == 0 and float(remain_hours) == 0:
                result_select_vc =  True
                vacation_request["vc_name"] = vc_name.replace(" ", "")
                cm.msg("p" , pr_rq.msg_re["pass_ad_select_vacation"])
                cm.msg("p" , "  +Vacation name : " + vc_name + " <Pass>" )
                break

            # If vacation is grant/regular , need to check there are enough days to request #
            else :
                if float(remain_days) >= hour_use :
                    result_select_vc = True
                    vacation_request["vc_name"] = vc_name.replace(" ", "")
                    cm.msg("p" , pr_rq.msg_re["pass_ad_select_vacation"])
                    cm.msg("p" , "  +Vacation name : " + vc_name + " <Pass>")
                    
                    break 
            i = i+1
        # Vacation does not have enough days to choose #
        if result_select_vc == False:
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_ad_no_vacation"])
        
        choose_vacation = xpath.par_choose_vacation(hour_use,result_select_vc,use_hour_unit,vacation_request,vc_name)
        return choose_vacation
    
    def request_all_day(request_date,approver):
        all                   = xpath.par_all_day()
        vacation_request      = xpath.par_vacation_request(request_date)
        vacation_request      = cm_rq.update_status_for_request(approver,vacation_request)
        choose_date           = cm_rq.choose_date_to_request(vacation_request,"all")
        cm_rq.all_choose_date(all,choose_date)
        cm_rq.select_approver()
        cm_rq.all_choose_vc(all,vacation_request,all_day)
        cm_rq.select_cc_enter_reason()
        return all

    def __init__(self,all_day):
        cm_rq.init(self,all_day)
       
class consecutive :
    
    def choose_vacation_to_request(vacation_request):
        cm.msg("n" ,"Vacation Name")
        result_select_vc = False
        use_hour_unit    = False
        total_vc = cm_rq.total_vacation()
        infor_vacation = cm_rq.check_use_hour_unit_half_day(total_vc)
        
        # Check there are any vacations using hour unit    #  
        # Change the way to check the number of leave days #
        i=1
        while i <= total_vc:
            if infor_vacation[i]["use_hour_unit"] == True:
                use_hour_unit = True
                break
            i = i+1

        # Check the remaining days of each vacation are enough to request #
        i = 1
        while i <= total_vc:
            time.sleep(1)
            hour_use = 2
            driver.find_element_by_css_selector( pr_rq.rq_vc["select_vacation"]).click()
            driver.find_element_by_xpath("//body/div[4]/div/div/div[" + str(i) +"]").click()
            vc_use_for_request = cm_fu.vacation_use_for_request()
            remain_days = vc_use_for_request["number_of_days"]
            remain_hours = vc_use_for_request["number_of_hours"]
            vc_name = vc_use_for_request["vacation_name"]

            cm.msg("n" ,"Vacation Name")
            # If vacation is other vacation can use this vacation to request #
            if float(remain_days) == 0 and float(remain_hours) == 0:
                result_select_vc = True
                vacation_request["vc_name"] = vc_name.replace(" ", "")
                cm.msg("p" , pr_rq.msg_re["pass_ad_select_vacation"])
                cm.msg("p" , "  +Vacation name : " +vc_name + "<Pass>" )
                break

            # If vacation is grant/regular , need to check there are enough days to request #
            else :
                if float(remain_days) >= hour_use :
                    result_select_vc = True
                    vacation_request["vc_name"] = vc_name.replace(" ", "")
                    cm.msg("p" , pr_rq.msg_re["pass_ad_select_vacation"])
                    cm.msg("p" , "  +Vacation name : " +vc_name + "<Pass>" )
                    break 
            i = i+1
            
        # Vacation does not have enough days to choose #
        if result_select_vc == False:
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_cs_no_vacation"])

        choose_vacation = xpath.par_choose_vacation(hour_use,result_select_vc,use_hour_unit,vacation_request,vc_name)
        return choose_vacation

    def vacation_consecutive(request_date,approver):
        vc_conse              = xpath.par_all_day()
        vacation_request      = xpath.par_vacation_request(request_date)
        vacation_request      = cm_rq.update_status_for_request(approver,vacation_request)
        choose_date           = cm_rq.choose_date_to_request(vacation_request,"vc_con")
        cm_rq.all_choose_date(vc_conse,choose_date)
        cm_rq.all_choose_vc(vc_conse,vacation_request,consecutive)
        return vc_conse

    def __init__(self,vc_conse):
        cm_rq.init(self,vc_conse) 

class half_day:
    def choose_vacation_to_request(vacation_request,type):
       
        cm.msg("n" ,"Vacation Name")
        use_half_day     = 0
        list_vc_use_half = []
        use_hour_unit    = False
        result_select_vc = False
        total_vc         = cm_rq.total_vacation()
        infor_vacation   = cm_rq.check_use_hour_unit_half_day(total_vc)
        
        # Check there are any vacations using hour unit 
        # Change the way to check the number of leave days 
        # Check there are any vacations using half day 
        if type == "am" :
            if total_vc == 0:
                cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_am_no_vacation"])

            else:
                for i in range(1,len(infor_vacation)):
                    if infor_vacation[i]["use_hour_unit"] == True:
                        use_hour_unit = True
                        break
                for i in range(1,len(infor_vacation)):
                    if infor_vacation[i]["use_half_day"] == True:
                        list_vc_use_half.append(infor_vacation[i])
                        use_half_day = use_half_day + 1

                if  use_half_day == 0:
                    cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_am_no_use"])
                else:
                    hour_use = cm_fu.hour_used(use_hour_unit,"am")
                    result_select_vc = cm_rq.select_vacation_use_hour_unit_half_day(total_vc,list_vc_use_half,hour_use,"half-day")
                    if  result_select_vc != False :
                        vacation_request["vc_name"] = result_select_vc.replace(" ", "")
                        cm.msg("p" , pr_rq.msg_re["pass_ad_select_vacation"])
                        cm.msg("p" ,"  +Vacation name  : " + result_select_vc + "<Pass>" )
                        
                        driver.find_element_by_xpath( pr_rq.rq_vc["radi_am"]).click()
                        '''
                        time_request = driver.find_element_by_xpath(rq_vc["selected_date"]).text
                        time_time_card = collect_clock_in_from_time_card(time_clock_in,"half-day",vc_date)
                    
                        
                        if time_request == time_time_card:
                            msg("p" ,msg_re["pass_time_clockin"])
                        else:
                            msg("f" ,msg_re["fail_time_clockin"])
                        '''
                    else:
                        cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_am_number_not_enough"])
        else :
            if total_vc == 0:
                cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_pm_no_vacation"])
            else:
                for i in range(1,len(infor_vacation)):
                    if infor_vacation[i]["use_hour_unit"] == True:
                        use_hour_unit = True
                        break
                for i in range(1,len(infor_vacation)):
                    if infor_vacation[i]["use_half_day"] == True:
                        list_vc_use_half.append(infor_vacation[i])
                        use_half_day = use_half_day+1

                if  use_half_day == 0:
                    cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_pm_no_use"])
                else:
                    hour_use = cm_fu.hour_used(use_hour_unit,"pm")
                    result_select_vc = cm_rq.select_vacation_use_hour_unit_half_day(total_vc,list_vc_use_half,hour_use,"half-day")
                    if  result_select_vc != False :
                        vacation_request["vc_name"] = result_select_vc.replace(" ", "")
                        cm.msg("p" , pr_rq.msg_re["pass_ad_select_vacation"])
                        cm.msg("p" ,"  +Vacation name  : " + result_select_vc + "<Pass>" )
                        driver.find_element_by_xpath( pr_rq.rq_vc["radi_pm"]).click()
                    else:
                        cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_am_number_not_enough"])

        choose_vacation = xpath.par_choose_vacation(hour_use,result_select_vc,use_hour_unit,vacation_request,result_select_vc)
        return choose_vacation

    def vacation_half(request_date,approver,type):
        vc_half               = xpath.par_all_day()
        vacation_request      = xpath.par_vacation_request(request_date)
        vacation_request      = cm_rq.update_status_for_request(approver,vacation_request)
        choose_date           = cm_rq.choose_date_to_request(vacation_request,type)
        cm_rq.all_choose_date(vc_half,choose_date)
        cm_rq.half_choose_vc(vc_half,vacation_request,type,half_day)
        return vc_half

    def __init__(self,vc_half):
        cm_rq.init(self,vc_half)  

class hour_unit:
    def choose_vacation_to_request(vacation_request):
        cm.msg("n" ,"Vacation Name")
        hour_use         = 2
        result_unit      = 0
        use_hour_unit    = False
        result_select_vc = False
        list_vc_use_hour_unit = []
        total_vc = cm_rq.total_vacation()
        infor_vacation = cm_rq.check_use_hour_unit_half_day(total_vc)
        

        if total_vc == 0:
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_ad_no_vacation"])
        else:
            for i in range(1,len(infor_vacation)):
                if infor_vacation[i]["use_hour_unit"] == True:
                    list_vc_use_hour_unit.append(infor_vacation[i])
                    result_unit = result_unit + 1
            if  result_unit == 0:
                cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_hu_no_use"])
            else:
                use_hour_unit = True
                result_select_vc = cm_rq.select_vacation_use_hour_unit_half_day(total_vc,list_vc_use_hour_unit,hour_use,"hour")
                if  result_select_vc != False :
                    cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_ad_select_vacation"])
                    cm.msg("p" ,"  +Vacation name  : " + result_select_vc + "<Pass>" )
                    vacation_request["vc_name"] =  result_select_vc.replace(" ","")
                    hour_selected =  cm_rq.select_hour_use_hour_unit()
                    vacation_request["vc_date"] =  str(vacation_request["vc_date"])+ hour_selected.replace("(","").replace(")","")
                
                else:
                    cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_hu_number_not_enough"])

        choose_vacation = xpath.par_choose_vacation(hour_use,result_select_vc,use_hour_unit,vacation_request,result_select_vc)
        return choose_vacation
    
    def vacation_hour_unit(request_date,approver,type):
        vc_hour               = xpath.par_all_day()
        vacation_request      = xpath.par_vacation_request(request_date)
        vacation_request      = cm_rq.update_status_for_request(approver,vacation_request)
        choose_date           = cm_rq.choose_date_to_request(vacation_request,type)
        cm_rq.all_choose_date(vc_hour,choose_date)
        cm_rq.all_choose_vc(vc_hour,vacation_request,hour_unit)
        return vc_hour

    def __init__(self,vc_hour):
        cm_rq.init(self,vc_hour)  


class used :

    def select_used_date_to_request_leave():
        i = 1 
        list_vc = cm_rq.get_vacation_date_and_status()
        request_date = cm_fu.vacation_date_is_used(list_vc)
        cm_fu.click_date(request_date)

        # Select date from find for request vacation # 
        driver.find_element_by_link_text("Request Vacation").click()
        cm.popup_time_card()
        current_month = driver.find_element_by_xpath(pr_rq.rq_vc["current_month"]).text[5:None] 
        if int(request_date.month) < 10 :
            request_month = "0" + str(request_date.month)
        else :
            request_month = request_date.month

        if str(current_month) == str(request_month) :
            result_click = cm_fu.click_date(request_date)
            if result_click == True:
                cm.msg("p" , pr_rq.msg_re["pass_select_date"])
                return request_date
            else:
                cm.msg("p" , pr_rq.msg_re["fail_select_date"])
                return False
        else:
            driver.find_element_by_xpath(pr_rq.rq_vc["icon_next_month"]).click()
            result_click = cm_fu.click_date(request_date)
            if result_click == True:
                cm.msg("p" , pr_rq.msg_re["pass_select_date"])
                return request_date
            else:
                cm.msg("p" , pr_rq.msg_re["fail_select_date"])

    def request_all_day(request_date,approver):
        all                   = xpath.par_all_day()
        vacation_request      = xpath.par_vacation_request(request_date)
        vacation_request      = cm_rq.update_status_for_request(approver,vacation_request)
        choose_date           = used.select_used_date_to_request_leave()
        cm_rq.all_choose_date(all,choose_date)
        cm_rq.all_choose_vc(all,vacation_request,all_day)
       
        return all

    def __init__(self,all_day):
        cm_rq.init(self,all_day)



class cancel():

    def cancel_request(oneday,use_hour_unit):

        try:
            i = 1
            time_cancel  = True
            check_number = False
            result_cancel= False
            result_find  = False
            list_status  = ["Request"]
            today        = xpath.today()
            number_before_cancel = cm_rq.available_vacation()
            driver.find_element_by_link_text("My Vacation Status").click()
            time.sleep(3)
            rows = driver.find_elements_by_xpath( pr_rq.rq_vc["list_request"])
            total_request = cm.total_data(rows)
            
            cm.msg("n" ,"Cancel Request")
            if cm.is_Displayed( pr_rq.rq_vc["check_list_re"]) == True :
                cm.xlsx( pr_rq.my_cancel["pass"] , pr_rq.msg_re["pass_cancel_no_request"])
            else:
                # Get all requests that can be used for cancel #
                time.sleep(3)
                list_request  = driver.find_elements_by_xpath(pr_rq.rq_vc["list_request"])
                total_request = cm.total_data(list_request)
                while i<= total_request:
                    infor_request = cm_rq.info_request_list(i)
                    if cm.is_Displayed(cm.xpath2( pr_rq.rq_vc["re_ic"],str(i-1),"')]")) == True : 
                        
                        # If vacation request is approved , check vacation date (> or = ) today => Can cancel #
                        if infor_request["status"] not in list_status :
                            time_cancel = cm_rq.time_comparison(infor_request["vacation_time"],today) 
                           
                        if time_cancel == True :
                            result_find = True
                            status_before_cancel = infor_request["status"]
                            request_to_cancel = infor_request
                            driver.find_element_by_xpath(cm.xpath2( pr_rq.rq_vc["re_ic"],str(i-1),"')]")).click()

                            if cm.is_Displayed( pr_rq.rq_vc["bt_cancel_request"]) == True :  
                                cm.msg("p" , pr_rq.msg_re["pass_cancel_click"])
                                driver.find_element_by_xpath( pr_rq.rq_vc["bt_cancel_request"]).click()
                                time.sleep(1)

                                if cm.is_Displayed( pr_rq.rq_vc["bt_cancel"]) == False : 
                                    cm.xlsx( pr_rq.my_cancel["pass"] , pr_rq.msg_re["pass_cancel"])
                                    result_cancel = True
                                else:
                                    cm.xlsx( pr_rq.my_cancel["fail"] , pr_rq.msg_re["fail_cancel"])
                                break 
                            else:
                                cm.xlsx( pr_rq.my_cancel["fail"] , pr_rq.msg_re["fail_cancel_click"])
                    i = i+1 
                

                # If the request is canceled successfully, check the request has changed status #
                if result_cancel == True:
                    i = 1
                    driver.find_element_by_link_text("My Vacation Status").click()
                    while i<= total_request:
                        infor_request = cm_rq.info_request_list(i)
                        find_request_canceled = cm_rq.two_requests_are_the_same(request_to_cancel,infor_request)

                        if find_request_canceled == True:
                            if status_before_cancel == "Request" :
                                if infor_request["status"] == "Canceled":
                                    cm.xlsx( pr_rq.my_cancel["pass"] , pr_rq.msg_re["pass_cancel_status_canceled"])
                                    check_number = True
                                    break
                                else:
                                    cm.xlsx( pr_rq.my_cancel["fail"] , pr_rq.msg_re["fail_cancel_status_canceled"])
                                    break
                            else:
                                # Status before is "Approved" ,"Approved[1/3],..."
                                if  status_before_cancel == "Approved" :
                                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, cm.xpath3("tr",str(i), pr_rq.rq_vc["ic_detail"],i-1, pr_rq.rq_vc["ic_detail1"])))).click()
                                    time.sleep(2)
                                    list_approver = driver.find_elements_by_xpath( pr_rq.rq_vc["content_vc_approver"])
                                    total_approver = cm.total_data(list_approver)
                                    if total_approver == 0:
                                        if infor_request["status"] == "Canceled":
                                            cm.xlsx( pr_rq.my_cancel["pass"] , pr_rq.msg_re["pass_cancel_status_canceled"])
                                            check_number = True
                                            break
                                            
                                        else:
                                            cm.xlsx( pr_rq.my_cancel["fail"] , pr_rq.msg_re["fail_cancel_status_canceled"])
                                            break

                                if infor_request["status"] == "User cancel":
                                    cm.xlsx( pr_rq.my_cancel["pass"] , pr_rq.msg_re["pass_cancel_status_usercancel"])
                                    break
                                else:
                                    cm.xlsx( pr_rq.my_cancel["fail"] , pr_rq.msg_re["fail_cancel_status_usercancel"])
                                    break
                            
                        i = i+1 
                if check_number == True :
                    hours_used = float(re.search(r'\d+',infor_request["use"]).group(0)) 
                    number_after_cancel = cm_rq.available_vacation()
                    hours = infor_request["use"].strip()
                    if infor_request["vacation_name"].rfind("~") > 0 :
                        vc_name = infor_request["vacation_name"].replace("\n","")
                        vc_name = vc_name[None: int(vc_name.rfind("~"))] + " ~ " +vc_name[ int(vc_name.rfind("~"))+1:None]
                    else :
                        vc_name = infor_request["vacation_name"]
                        
                    if hours == "1D" or hours == "2D":
                        type_request = "all"
                    elif hours == "0.5D" or hours == "4H":
                        type_request = "half_day"
                    else:
                        type_request = "hour"

                    cm_rq.check_number_of_days_cancel(number_before_cancel,number_after_cancel,hours_used,vc_name,oneday,use_hour_unit,type_request)
                
                if result_find == False:
                    cm.xlsx( pr_rq.my_cancel["pass"] , pr_rq.msg_re["pass_cancel_no_request"])
        except:
            driver.find_element_by_link_text("My Vacation Status").click()

class filter():
     
    def xpath_name(type,j):
        if type == "cc" :
            xpath_name = xpath.filter_name_cc(j)
        else :
            xpath_name = xpath.filter_name_my(j)
        return xpath_name

    def list_status_before(type):
        # Get all status from list request vacation #
        i=1
        list_status_before = []
        time.sleep(3)
        driver.find_element_by_xpath(pr_mn.mn_pro["ic_to_end_page"]).click()
        end_page_text = driver.find_element_by_xpath(pr_mn.mn_pro["page_current"]).text
        total_page=int(end_page_text)
        driver.find_element_by_xpath(pr_mn.mn_pro["ic_to_first_page"]).click()
        time.sleep(1)
        while i <= total_page:
            if i == total_page :
                time.sleep(3)
                j = 1
                total_request = driver.find_elements_by_xpath(pr_mn.mn_pro["list_re_vc"])
                while j<= len(total_request):
                    xpath_name = filter.xpath_name(type,j)
                    status_name = driver.find_element_by_xpath(xpath_name).text
                    list_status_before.append(status_name)
                    j = j+1
            else:
                j = 1
                while j <= 20:
                    xpath_name = filter.xpath_name(type,j)
                    status_name = driver.find_element_by_xpath(xpath_name).text
                    list_status_before.append(status_name)
                    j = j+1
            time.sleep(2)
            total_ic = driver.find_elements_by_xpath(pr_mn.mn_pro["total_ic"])
            driver.find_element_by_xpath(xpath.filter_page(total_ic)).click()
            i=i+1
        return list_status_before
    
    def filter_status(sta_name_to_filter,select_status,submenu,list_status_before):
        j = 1
        total_status = 0
        result_filter = True
        list_status_progressing = ["Request","Approved","Canceled"]
        list_status_Completed = ["Approved","Canceled"]

        if submenu == "myvacation":
            xpath_status = pr_rq.rq_vc["fi_re_status"]
            param_ex = pr_rq.param_excel_re["filter1"]
        if submenu == "viewcc" :
            xpath_status = pr_rq.rq_vc["re_status_cc"]
            param_ex = pr_rq.param_excel_re["filter2"]

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR,pr_rq.rq_vc["open_filter"]))).click()
        driver.find_element_by_xpath(pr_rq.rq_vc[select_status]).click()

        # FILTER STATUS IS REQUEST #
        if sta_name_to_filter == "Request":
            # Count status is request from all status #
            for i in list_status_before:
                if i == sta_name_to_filter:
                    total_status = total_status + 1

            # Filtered list is empty #
            if cm.is_Displayed(pr_rq.rq_vc["text_list_ap1"]) == True:
                if total_status == 0 :
                    cm.msg("p",xpath.filter("p",sta_name_to_filter))
                   
                else:
                    result_filter=False
                    cm.xlsx(param_ex["fail"] ,xpath.filter("f",sta_name_to_filter))

            # Filtered list is not empty => Check filter list contains other status  #
            else:
                time.sleep(3)
                rows = driver.find_elements_by_xpath(pr_rq.rq_vc["list_request"])
                total_request = cm.total_data(rows)
                while j<= total_request :
                    status = driver.find_element_by_xpath("//tr["+str(j)+xpath_status).text
                    if status != sta_name_to_filter:
                        cm.xlsx(param_ex["fail"] ,xpath.filter("f",sta_name_to_filter))
                        result_filter = False
                        break
                    j=j+1
                if result_filter == True:
                    cm.msg("p",xpath.filter("p",sta_name_to_filter))
                

        # FILTER STATUS IS PROGRESSING #
        elif sta_name_to_filter == "Progressing":
            # Count status is request from all status #
            for i in list_status_before:
                if i not in  list_status_progressing:
                    total_status= total_status + 1
            
            # Filtered list is empty #
            if cm.is_Displayed(pr_rq.rq_vc["text_list_ap1"]) == True:
                if total_status == 0 :
                    cm.msg("p",xpath.filter("p",sta_name_to_filter))
                else:
                    result_filter=False
                    cm.xlsx(param_ex["fail"] ,xpath.filter("f",sta_name_to_filter))

            # Filtered list is not empty => Check filter list contains other status  #
            else:
                time.sleep(3)
                rows = driver.find_elements_by_xpath(pr_rq.rq_vc["list_request"])
                total_request = cm.total_data(rows)
                while j <= total_request :
                    status = driver.find_element_by_xpath("//tr["+str(j)+xpath_status).text
                    if status in  list_status_progressing:
                        cm.xlsx(param_ex["fail"] ,xpath.filter("f",sta_name_to_filter))
                        result_filter = False
                        break
                    j = j+1
                if result_filter == True:
                    cm.msg("p",xpath.filter("p",sta_name_to_filter))
                    

        # FILTER STATUS IS COMPLETED #
        else:
            # Count status is request from all status #
            for i in list_status_before:
                if i in  list_status_Completed:
                    total_status= total_status + 1

            # Filtered list is empty #
            if cm.is_Displayed(pr_rq.rq_vc["text_list_ap1"]) == bool(True):
                if total_status == 0 :
                    cm.msg("p",xpath.filter("p",sta_name_to_filter))
                else:
                    result_filter=False
                    cm.xlsx(param_ex["fail"] ,xpath.filter("f",sta_name_to_filter))

            # Filtered list is not empty => Check filter list contains other status  #
            else:
                time.sleep(3)
                rows = driver.find_elements_by_xpath(pr_rq.rq_vc["list_request"])
                total_request = cm.total_data(rows)
                while j <= total_request :
                    status = driver.find_element_by_xpath("//tr["+str(j)+xpath_status).text
                    if status not in list_status_Completed:
                        cm.xlsx(param_ex["fail"] ,xpath.filter("f",sta_name_to_filter))
                        result_filter = False
                        break
                    j = j+1
                if result_filter == True:
                    cm.msg("p",xpath.filter("p",sta_name_to_filter))

        return result_filter 

class view():
    def view_detail_request_at_calendar(param):
        time.sleep(5)
        requests_list = driver.find_elements_by_xpath(pr_rq.rq_vc["detail_request"])
        if len(requests_list) > 0:
            requests_list[0].click()
            cm.xlsx(param["pass"] ,pr_rq.msg_re["pass_view"])
        else:
            cm.xlsx(param["fail"] ,pr_rq.msg_re["pass_no_view"])

