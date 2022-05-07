
from kq_hr_login import driver
from kq_hr_lib_common import cm,time,Keys,relativedelta,EC,WebDriverWait,By,data ,datetime
from kq_hr_param import pr_mn


#param_excel = data["testcase_result"]["manager_pro"]

def filter_status(sta_name_to_filter,select_status,list_status_before):
    list_status_progressing = ["Request","Approved","Canceled","Waiting"]
    list_status_Completed = ["Approved","Canceled"]
    result_filter = True
    total_status = 0
   


    time.sleep(3)
    driver.find_element_by_xpath(pr_mn.mn_pro["open_fil_status"]).click()
    driver.find_element_by_xpath(pr_mn.mn_pro[select_status]).click()

    # FILTER STATUS IS REQUEST #
    if sta_name_to_filter == "Request" or sta_name_to_filter =="Waiting":
        # Count status is request from all status #
        for i in list_status_before:
            if i == sta_name_to_filter:
                total_status= total_status + 1

        # Filtered list is empty #
        if cm.is_Displayed(pr_mn.mn_pro["list_request"]) == True:
            if total_status == 0 :
                cm.msg("p","  +Filter by status is "+ sta_name_to_filter + " <Pass>")
            else:
                result_filter=False
                cm.xlsx(pr_mn.mn_filter["fail"] ,"  +Filter by status is "+ sta_name_to_filter + " <Fail>")

        # Filtered list is not empty => Check filter list contains other status  #
        else:
            j=1
            time.sleep(3)
            total_request=driver.find_elements_by_xpath(pr_mn.mn_pro["list_re_vc"])
            while j<= len(total_request):
                status_name=driver.find_element_by_xpath(pr_mn.vc_ap["cl_status1"]+str(j)+pr_mn.vc_ap["cl_status2"]).text
                if status_name != sta_name_to_filter:
                    result_filter=False
                    cm.xlsx(pr_mn.mn_filter["fail"] ,"  +Filter by status is "+ sta_name_to_filter + " <Fail>")
                    break
                j=j+1
            if result_filter == True:
                cm.msg("p","  +Filter by status is "+ sta_name_to_filter + " <Pass>")
     
    # FILTER STATUS IS PROGRESSING #
    elif sta_name_to_filter == "Progressing":
        # Count status is request from all status #
        for i in list_status_before:
           if i not in  list_status_progressing:
               total_status= total_status + 1
        # Filtered list is empty #
        if cm.is_Displayed(pr_mn.mn_pro["list_request"]) == True:
            if total_status == 0 :
                cm.msg("p","  +Filter by status is "+ sta_name_to_filter + " <Pass>")
            else:
                result_filter=False
                cm.xlsx(pr_mn.mn_filter["fail"] ,"  +Filter by status is "+ sta_name_to_filter + " <Fail>")

        # Filtered list is not empty => Check filter list contains other status  #
        else:
            time.sleep(3)
            j=1
            rows=driver.find_elements_by_xpath(data["rq_vc"]["list_request"])
            total_request = cm.total_data(rows)
            while j<= total_request :
                status_name=driver.find_element_by_xpath(pr_mn.vc_ap["cl_status1"]+str(j)+pr_mn.vc_ap["cl_status2"]).text
                if status_name in  list_status_progressing:
                    result_filter=False
                    cm.xlsx(pr_mn.mn_filter["fail"] ,"  +Filter by status is "+ sta_name_to_filter + " <Fail>")
                    break
                j=j+1
            if result_filter== True:
                cm.msg("p","  +Filter by status is "+ sta_name_to_filter + " <Pass>")
                
    
    # FILTER STATUS IS COMPLETED #
    else:
        # Count status is request from all status #
        for i in list_status_before:
           if i in  list_status_Completed:
               total_status= total_status + 1

        # Filtered list is empty #
        if cm.is_Displayed(pr_mn.mn_pro["list_request"]) == True:
            if total_status == 0 :
                cm.msg("p","  +Filter by status is "+ sta_name_to_filter + " <Pass>")
            else:
                result_filter=False
                cm.xlsx(pr_mn.mn_filter["fail"] ,"  +Filter by status is "+ sta_name_to_filter + " <Fail>")

        # Filtered list is not empty => Check filter list contains other status  #
        else:
            time.sleep(3)
            rows=driver.find_elements_by_xpath(data["rq_vc"]["list_request"])
            total_request = cm.total_data(rows)
            j=1
            while j<= total_request :
                status_name=driver.find_element_by_xpath(pr_mn.vc_ap["cl_status1"]+str(j)+pr_mn.vc_ap["cl_status2"]).text
                if status_name not in list_status_Completed:
                    result_filter=False
                    cm.xlsx(pr_mn.mn_filter["fail"] ,"  +Filter by status is "+ sta_name_to_filter + " <Fail>")
                    break
                j=j+1
            if result_filter == True:
                cm.msg("p","  +Filter by status is "+ sta_name_to_filter + " <Pass>")

    
    return result_filter 
    
def select_user_from_depart(type):
    # Select user from org #
    if type == "adjust" :
        driver.find_element_by_xpath(data["sm_vacation_adjust"]).click()
        time.sleep(3)
        list_depart = driver.find_elements_by_xpath(pr_mn.mn_pro["list_depart"])
        print("list_depart :", cm.total_data(list_depart))
        for i in range(1,len(list_depart)):
            time.sleep(1)
            depart_has_user = cm.is_Displayed(pr_mn.mn_pro["depart_name"]+str(i)+pr_mn.mn_pro["depart_name1"]) 
            if depart_has_user == True :
                driver.find_element_by_xpath(pr_mn.mn_pro["depart_name"]+str(i)+pr_mn.mn_pro["depart_name1"]).click()
                time.sleep(1)
                total_user= driver.find_elements_by_xpath(pr_mn.mn_pro["list_user"])
                for i in range(1,len(total_user)+1):
                    is_user = cm.is_Displayed(pr_mn.mn_pro["sl_user"]+str(i)+pr_mn.mn_pro["sl_user1"]) 
                    if is_user == True:
                        user_name = driver.find_element_by_xpath(pr_mn.mn_pro["user_name"] + str(i) + pr_mn.mn_pro["user_name1"])
                        user_name.click()
                        return True  
    else :
        time.sleep(3)
        list_depart = driver.find_elements_by_xpath(pr_mn.mn_pro["list_depart"])
        for i in range(1,len(list_depart)):
            time.sleep(1)
            depart_has_user = cm.is_Displayed(pr_mn.mn_pro["is_depart"]+str(i)+pr_mn.mn_pro["is_depart1"]) 
            if depart_has_user == True :
                driver.find_element_by_xpath(pr_mn.mn_pro["is_depart"]+str(i)+pr_mn.mn_pro["is_depart1"]).click()
                time.sleep(1)
                total_user= driver.find_elements_by_xpath(pr_mn.mn_pro["list_user"])
                for i in range(1,len(total_user)+1):
                    is_user = cm.is_Displayed(pr_mn.mn_pro["sl_user"]+str(i)+pr_mn.mn_pro["sl_user1"]) 
                    if is_user == True:
                        user_name = driver.find_element_by_xpath(pr_mn.mn_pro["user_name"] + str(i) + pr_mn.mn_pro["user_name1"])
                        user_name.click()
                        return True  

    return False 

def search_user_and_select_user_from_org():
    try:
        # Total user/department before search and  first username/first departmentname  #
        time.sleep(4)
        total_department_before=driver.find_elements_by_xpath(pr_mn.mn_pro["list_user_org"])
        first_username=driver.find_element_by_xpath(pr_mn.mn_pro["choose_user_0"]).text

        # Search user #
        ip_search_user=driver.find_element_by_xpath(pr_mn.mn_pro["ip_search_user"])
        ip_search_user.send_keys(pr_mn.user)
        ip_search_user.send_keys(Keys.RETURN)
        if ip_search_user.get_attribute('value') == pr_mn.user :
            cm.msg("p",pr_mn.msg_mn["pass_search_enter"])

            if cm.is_Displayed(pr_mn.mn_pro["no_data"]) == False:
                time.sleep(3)
                total_department_after=driver.find_elements_by_xpath(pr_mn.mn_pro["list_user_org"])
                user_firt_search=driver.find_element_by_xpath(pr_mn.mn_pro["choose_user_0"])
                first_username_after=user_firt_search.text
                if len(total_department_after)==len(total_department_before):
                    if first_username==first_username_after:
                        cm.xlsx(pr_mn.mn_search["fail"] ,pr_mn.msg_mn["fail_search"])
                    else:
                        cm.xlsx(pr_mn.mn_search["pass"] ,pr_mn.msg_mn["pass_search"])

                else:
                    cm.xlsx(pr_mn.mn_search["pass"] ,pr_mn.msg_mn["pass_search"])
                    user_firt_search.click()
                    if cm.is_Displayed(pr_mn.mn_pro["sl_user_grant"]) ==False :
                        cm.xlsx(pr_mn.mn_search["pass"] ,pr_mn.msg_mn["pass_search_choose_user"])
            else:
                cm.xlsx(pr_mn.mn_search["pass"] ,pr_mn.msg_mn["pass_search_not_exist"])
        else:
            cm.xlsx(pr_mn.mn_search["fail"] ,pr_mn.msg_mn["fail_search_enter"])

        # Select user from org #
        
        select_user_from_depart("adjust")
        if cm.is_Displayed(pr_mn.mn_pro["sl_user_grant"]) ==False :
            cm.xlsx(pr_mn.mn_ogr["pass"] ,pr_mn.msg_mn["pass_org_choose"])
            return True
        else:
            cm.xlsx(pr_mn.mn_ogr["fail"] ,pr_mn.msg_mn["fail_org_choose"])
            return False
    except:
        driver.find_element_by_xpath(data["sm_vacation_adjust"]).click()

def sm_vc_ad_grant_vacation():
    # Grant Vacation #
    try :
        nb_days="10"
        nb_hour="4"
        
        select_user = select_user_from_depart("adjust")
        driver.find_element_by_css_selector(pr_mn.mn_pro["sort_search"]).click()
        if select_user == True :
            # Select vacation name to grant #
            cm.msg("p","  [Set up number of days to grant]")
            time.sleep(2)
            result_select_vc=False 
            driver.find_element_by_css_selector(pr_mn.mn_pro["select_vc"]).click()
            time.sleep(2)
            vacation_name=driver.find_element_by_xpath(pr_mn.mn_pro["choose_vc"])
            if vacation_name.text == "No options" :
                cm.xlsx(pr_mn.mn_grant["pass"] ,pr_mn.msg_mn["pass_no_grant"])
            else:
                vacation_name.click()
                result_select_vc=True
                cm.msg("p",pr_mn.msg_mn["pass_grant_vc_name"])
                

        
            # Enter days to grant #
            time.sleep(2)
            days=driver.find_element_by_xpath(pr_mn.mn_pro["ip_day"])
            days.clear()
            days.send_keys(nb_days)
            days.send_keys(Keys.RETURN)
        
            if days.get_attribute('value')==nb_days :
                cm.msg("p",pr_mn.msg_mn["pass_grant_enter"])
                
            else:
                cm.xlsx(pr_mn.mn_grant["fail"] ,pr_mn.msg_mn["fail_grant_enter"])
            
            
            # Enter hour #
            try:
                hour=driver.find_element_by_xpath(pr_mn.mn_pro["ip_hour"])
                hour.clear()
                hour.send_keys(nb_hour)
                hour.send_keys(Keys.RETURN)
                if hour.get_attribute('value')==nb_hour :
                    cm.xlsx(pr_mn.mn_grant["pass"] ,pr_mn.msg_mn["pass_grant_enter_hours"])
                else:
                    cm.xlsx(pr_mn.mn_grant["fail"] ,pr_mn.msg_mn["fail_grant_enter_hours"])
            except:
                cm.msg("p",pr_mn.msg_mn["fail_grant_use_unit"])
                

            
            # Enter memo #

            time.sleep(1)
            driver.find_element_by_css_selector(pr_mn.mn_pro["ic_memo_grant"]).click()

            memo=driver.find_element_by_xpath(pr_mn.mn_pro["ip_memo"])
            memo.clear()
            memo.send_keys(pr_mn.msg_mn["grant_memo"])
            memo.send_keys(Keys.RETURN)
            time.sleep(1)
            if memo.get_attribute('value').strip()==pr_mn.msg_mn["grant_memo"] :
                cm.msg("p",pr_mn.msg_mn["pass_grant_memo"])
            else:
                cm.xlsx(pr_mn.mn_grant["fail"] ,pr_mn.msg_mn["fail_grant_memo"])

            driver.find_element_by_xpath(pr_mn.mn_pro["bt_save_memo"]).click()
            if cm.is_Displayed(pr_mn.mn_pro["bt_cancel_memo"]) ==False:
                cm.msg("p",pr_mn.msg_mn["pass_grant_save"])
            else:
                cm.xlsx(pr_mn.mn_grant["fail"] ,pr_mn.msg_mn["fail_grant_save"])

            # Save Grant #
            if result_select_vc == True:
                driver.find_element_by_xpath(pr_mn.mn_pro["bt_be_grant"]).click()
                time.sleep(3)
                if cm.is_Displayed(pr_mn.mn_pro["af_be_grant"]) == True :
                    cm.msg("p","  +Click on Grant button <Pass>")
                    driver.find_element_by_xpath(pr_mn.mn_pro["af_be_grant"]).click()
                    time.sleep(3)
                    if cm.is_Displayed(pr_mn.mn_pro["bt_close"]) ==False :
                        cm.xlsx(pr_mn.mn_grant["pass"] ,pr_mn.msg_mn["pass_grant"])
                    else:
                        cm.xlsx(pr_mn.mn_grant["fail"] ,pr_mn.msg_mn["fail_grant"])
                    
                else:
                    cm.xlsx(pr_mn.mn_grant["fail"] ,pr_mn.msg_mn["fail_grant_click"])
            else:
                cm.xlsx(pr_mn.mn_grant["pass"] ,pr_mn.msg_mn['pass_conditions_grant'])
    except:
        driver.find_element_by_xpath(data["sm_vacation_adjust"]).click()

def next_date(request_date):
    if int(data["month"][str(request_date.month)]) == request_date.day:
        request_date= request_date + relativedelta(month=request_date.month+1) + relativedelta(day=1)
   
    elif request_date.weekday() == 5 :
        request_date= request_date + relativedelta(day=request_date.day +2)
        if int(data["month"][str(request_date.month)]) == request_date.day:
            request_date= request_date + relativedelta(month=request_date.month+1) + relativedelta(day=1)
        
    elif  request_date.weekday() == 6 :
        request_date= request_date + relativedelta(day=request_date.day +1)
        if int(data["month"][str(request_date.month)]) == request_date.day:
            request_date= request_date + relativedelta(month=request_date.month+1) + relativedelta(day=1)
    else :
        request_date= request_date + relativedelta(day=request_date.day +1)
   
    return request_date

def split_date_from_continuous_date(continuous_date,date_used):
    if continuous_date.rfind("~") > 0 :
        start_date = continuous_date[None: int(continuous_date.rfind("~"))]
        start_date  =  datetime.datetime.strptime(start_date , '%Y-%m-%d').date()
        end_date = continuous_date[int(continuous_date.rfind("~"))+1: None]
        end_date  =  datetime.datetime.strptime(end_date , '%Y-%m-%d').date()
        next_date_1 = start_date
        while next_date_1 !=  end_date :
            date_used.append(str(next_date_1))
            if start_date == end_date :
                break
            next_date_1 = next_date(next_date_1)
        date_used.append(str(end_date))
    else :
        date_used.append(continuous_date)

def choose_date():
    i = 1
    date_used = []
    
    # Go to my vacation to take used days # 
    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.LINK_TEXT,"My Vacation Status"))).click()
    time.sleep(3)
    rows=driver.find_elements_by_xpath(pr_mn.rq_vc["list_request"])
    total_request = cm.total_data(rows)
    while i <= total_request:
        if i == 1:
            if cm.is_Displayed(pr_mn.rq_vc["check_list_re"]) == True :
               break
            else:
                date=driver.find_element_by_xpath(cm.xpath("tr",str(i),pr_mn.rq_vc["request"])).text
                split_date_from_continuous_date(date,date_used)
        else:
            date=driver.find_element_by_xpath(cm.xpath("tr",str(i),pr_mn.rq_vc["request"])).text
            split_date_from_continuous_date(date,date_used)
        i=i+1

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
        
        while str(request_date) in date_used :
            request_date = next_date(request_date)
            if request_date.weekday() == 5 :
                request_date = request_date + relativedelta(day=request_date.day +2)
            if  request_date.weekday() == 6 :
                request_date = request_date + relativedelta(day=request_date.day +1)
            if int(data["month"][str(request_date.month)]) == request_date.day:
                request_date = request_date + relativedelta(month = request_date.month+1) + relativedelta(day = 1)
    return request_date
   
def click_date(request_date) :
   
    current_month = driver.find_element_by_xpath(pr_mn.mn_pro["month_name"]).text
    current_month = current_month[5:7]
    while str(request_date.month) != current_month :
        driver.find_element_by_xpath(pr_mn.mn_pro["next_month"]).click()
        current_month = driver.find_element_by_xpath(pr_mn.mn_pro["month_name"]).text
        current_month = current_month[5:7]

    if request_date.day < 25 :
        for i in range(2,7) :
            for j in range(2,7):
                date_at_calendar = driver.find_element_by_xpath("//tr["+str(i)+"]/td["+str(j)+"]/span")
                if str(date_at_calendar.text) == str(request_date.day):
                    date_at_calendar.click()
                    selected_date = driver.find_element_by_xpath(pr_mn.mn_pro["current_month"]).text
                    selected_date = selected_date[None: int(selected_date.rfind("("))]
                    selected_date = selected_date[5:7]
                    if selected_date.rfind("(") > 0:
                        selected_date = selected_date[None: int(selected_date.rfind("("))]
                    else:
                        selected_date = selected_date[12: None].replace(" ", "")
                    return True
        return False

    else :
        for i in range(2,7) :
            for j in range(2,7):
                date_at_calendar = driver.find_element_by_xpath("//tr["+str(i)+"]/td["+str(j)+"]/span")
                if str(date_at_calendar.text) == str(request_date.day):
                    date_at_calendar.click()
                    selected_date = driver.find_element_by_xpath(pr_mn.mn_pro["current_month"]).text
                    selected_date = selected_date[None: int(selected_date.rfind("("))]
                    selected_date = selected_date[5:7]
                    if selected_date.rfind("(") > 0:
                        selected_date = selected_date[None: int(selected_date.rfind("("))]
                    else:
                        selected_date = selected_date[12: None].replace(" ", "")

                    if str(request_date) == selected_date:
                        return True
                    else:
                        date_at_calendar.click()
                        
        return False

def sm_vc_ad_adjust_vacation():
    # Adjust Vacation #
    try:
        number_adjust = 2
        driver.find_element_by_xpath(data["sm_vacation_adjust"]).click()
        select_user = select_user_from_depart("adjust")
        driver.find_element_by_css_selector(pr_mn.mn_pro["sort_search"]).click()
        if select_user == True :
            driver.find_element_by_xpath(pr_mn.mn_pro["ra_adjust"]).click()
            if cm.is_Displayed(pr_mn.mn_pro["bt_adjust"]) == True :
                cm.msg("p",pr_mn.msg_mn["pass_adjust_radio"])
                
                # Select vacation name to Adjust #
                result_select_vc=False 
                time.sleep(2)
                vacation_name=driver.find_element_by_xpath(pr_mn.mn_pro["vc_name_adjust"]).text
                if len(vacation_name) !=0:
                    result_select_vc=True
                    ip_hour=driver.find_element_by_xpath(pr_mn.mn_pro["ip_hour_adjust"])
                    available_hours=int(ip_hour.get_attribute('value'))
                    
                    hour_used=driver.find_element_by_xpath(pr_mn.mn_pro["used_adjust"]).text
                    if hour_used=="-":
                        hour_used=0
                    
                    hour_remain=driver.find_element_by_xpath(pr_mn.mn_pro["remain_adjust"]).text
                    if hour_remain.rfind("D") >0 :
                        hour_remain=int(hour_remain[None: int(hour_remain.rfind("D"))])
                    else:
                        hour_remain=0

                    ip_hour.clear()
                    ip_hour.send_keys(number_adjust+available_hours)
                    available_hours1=int(ip_hour.get_attribute('value'))
                    if available_hours1 == number_adjust+available_hours:
                        cm.msg("p",pr_mn.msg_mn["pass_adjust_enter_hours"])
                        
                    else:
                        cm.xlsx(pr_mn.mn_adjust["fail"] ,pr_mn.msg_mn["fail_adjust_enter_hours"])
                    
                    hour_remain1=driver.find_element_by_xpath(pr_mn.mn_pro["remain_adjust"]).text
                    if hour_remain1.rfind("D") >0 :
                        hour_remain1=int(hour_remain1[None: int(hour_remain1.rfind("D"))])
                    else:
                        hour_remain1=0
                    if hour_remain1 == hour_remain+number_adjust:
                        cm.msg("p",pr_mn.msg_mn["pass_adjust_remain_column"])
                    else:
                        cm.xlsx(pr_mn.mn_adjust["fail"] ,pr_mn.msg_mn["fail_adjust_remain_column"])
                    if result_select_vc == True:
                        driver.find_element_by_xpath(pr_mn.mn_pro["bt_adjust"]).click()
                        if cm.is_Displayed(pr_mn.mn_pro["bt_close1"]) == True : 
                            cm.msg("p",pr_mn.msg_mn["pass_adjust_click"])

                            driver.find_element_by_xpath(pr_mn.mn_pro["bt_adjust1"]).click()
                            if cm.is_Displayed(pr_mn.mn_pro["bt_close"]) ==False : 
                                cm.xlsx(pr_mn.mn_adjust["pass"] ,pr_mn.msg_mn["pass_adjust"])
                            else:
                                cm.xlsx(pr_mn.mn_adjust["fail"] ,pr_mn.msg_mn["fail_adjust"])
                        else:
                            cm.xlsx(pr_mn.mn_adjust["fail"] ,pr_mn.msg_mn["fail_adjust_click"])
                    
                else:
                    cm.xlsx(pr_mn.mn_adjust["pass"] ,pr_mn.msg_mn["pass_no_adjust"])
            
            else:
                cm.xlsx(pr_mn.mn_adjust["fail"] ,pr_mn.msg_mn["fail_adjust_radio"])
    except:
        driver.find_element_by_xpath(data["sm_vacation_adjust"]).click()

def sm_vc_ad_register_usage_history():

    request_date = choose_date()
    print(request_date)
    # Register Usage History #
    user = data["user"]
    result_vc = False
    

    # Search user #
    driver.find_element_by_xpath(data["sm_vacation_adjust"]).click()
    driver.find_element_by_xpath(pr_mn.mn_pro["tab_vc_adjust"]).click()

    # Total user/department before search and  first username/first departmentname  

    ip_search_user=driver.find_element_by_xpath(pr_mn.mn_pro["ip_search_user"])
    ip_search_user.send_keys(user)
    ip_search_user.send_keys(Keys.RETURN)
    if ip_search_user.get_attribute('value') == user :
        cm.msg("p",pr_mn.msg_mn["pass_search_enter"])

        if cm.is_Displayed(pr_mn.mn_pro["no_data"]) == False:
            driver.find_element_by_xpath(pr_mn.mn_pro["ch_user"]).click()
            driver.find_element_by_xpath(pr_mn.mn_pro["ra_register"]).click()
            if cm.is_Displayed(pr_mn.mn_pro["bt_upload"]) == True :
                result_vc== True
                cm.xlsx(pr_mn.mn_register["pass"] ,pr_mn.msg_mn["pass_register_radio"])

                vc_name=driver.find_element_by_xpath(pr_mn.mn_pro["register_vc_name"]).text
                if len(vc_name)!=0:
                    remain=driver.find_element_by_xpath(pr_mn.mn_pro["register_remain"]).text
                    remain=int(remain.replace("D", ""))
                else:
                    cm.xlsx(pr_mn.mn_register["pass"] ,pr_mn.msg_mn["pass_no_register"])
                driver.find_element_by_xpath(pr_mn.mn_pro["add_date"]).click()
                if cm.is_Displayed(pr_mn.mn_pro["bt_add_date"]) == True : 
                    cm.msg("p",pr_mn.msg_mn["pass_register_date"])
                    driver.find_element_by_xpath(pr_mn.mn_pro["bt_add_date"]).click()
                    driver.find_element_by_xpath(pr_mn.mn_pro["date_and_time"]).click()
                    click_date(request_date)
                    driver.find_element_by_xpath(pr_mn.mn_pro["bt_selected"]).click()
                    driver.find_element_by_xpath(pr_mn.mn_pro["bt_save"]).click()
                    if cm.is_Displayed(pr_mn.mn_pro["add_date"]) == True :  
                        cm.msg("p",pr_mn.msg_mn["pass_save_date"])
                        driver.find_element_by_xpath(pr_mn.mn_pro["bt_upload"]).click()
                        if cm.is_Displayed(pr_mn.mn_pro["bt_close"]) == True : 
                            cm.msg("p",pr_mn.msg_mn["pass_upload"])
                            driver.find_element_by_xpath(pr_mn.mn_pro["bt_upda"]).click()
                            if cm.is_Displayed(pr_mn.mn_pro["add_date"]) == True : 
                                cm.msg("p",pr_mn.msg_mn["pass_register"])
                            else:
                                cm.xlsx(pr_mn.mn_register["fail"] ,pr_mn.msg_mn["fail_register"])

                        else:
                            cm.xlsx(pr_mn.mn_register["fail"] ,pr_mn.msg_mn["fail_upload"])

                    else :
                        cm.xlsx(pr_mn.mn_register["fail"] ,pr_mn.msg_mn["fail_save_date"])


                else :
                    cm.xlsx(pr_mn.mn_register["fail"] ,pr_mn.msg_mn["fail_register_date"])
                    


        else:
            cm.xlsx(pr_mn.mn_register["fail"] ,pr_mn.msg_mn["fail_register_radio"])
            
def sm_vc_ap_view_detail_request():
    # View detail request #
    try:
        if cm.is_Displayed(pr_mn.mn_pro["list_request"]) ==False :
            driver.find_element_by_xpath(pr_mn.mn_pro["ic_detail"]).click()
            if cm.is_Displayed(pr_mn.mn_pro["check_view"]) == True :
                cm.msg("p",pr_mn.msg_mn["pass_view_icon"])

                vc_date_title=driver.find_element_by_xpath(pr_mn.mn_pro["content_request"]).text
                vc_date_content=driver.find_element_by_xpath(pr_mn.mn_pro["content_request1"]).text 
                if vc_date_title.strip() == "Vacation Date" and len(vc_date_content.strip()) != 0 :
                    cm.msg("p",pr_mn.msg_mn["pass_view_date"])
                    result_vc_date=True
                else:
                    cm.xlsx(pr_mn.mn_detail_re["fail"] ,pr_mn.msg_mn["fail_view_date"])

                use_title=driver.find_element_by_xpath(pr_mn.mn_pro["content_use"]).text
                use_content=driver.find_element_by_xpath(pr_mn.mn_pro["content_use1"]).text 
                if use_title.strip() == "Use" and len(use_content.strip()) != 0 :
                    cm.msg("p",pr_mn.msg_mn["pass_view_number_used"])
                    result_number_use=True
                else:
                    cm.xlsx(pr_mn.mn_detail_re["fail"] ,pr_mn.msg_mn["fail_view_number_used"])
                
            
                request_date_title=driver.find_element_by_xpath(pr_mn.mn_pro["content_re_date"]).text
                request_date_content=driver.find_element_by_xpath(pr_mn.mn_pro["content_re_date1"]).text
                if request_date_title.strip() =="Request date" and len(request_date_content.strip()) !=0:
                    cm.msg("p",pr_mn.msg_mn["pass_view_rq_date"])
                    result_date=True
                else:
                    cm.xlsx(pr_mn.mn_detail_re["fail"] ,pr_mn.msg_mn["fail_view_rq_date"])
                

                if result_date== True and result_number_use==True and result_vc_date== True:
                    cm.xlsx(pr_mn.mn_detail_re["pass"] ,pr_mn.msg_mn["pass_view"])
                else:
                    cm.xlsx(pr_mn.mn_detail_re["fail"] ,pr_mn.msg_mn["fail_view"])

                driver.find_element_by_link_text("Vacation Approve").click()
            else:
                cm.xlsx(pr_mn.mn_detail_re["fail"] ,pr_mn.msg_mn['fail_view_icon'])
        else:
            cm.xlsx(pr_mn.mn_detail_re["pass"] ,pr_mn.msg_mn["pass_no_view"])
    except:
        driver.find_element_by_link_text("Vacation Approve").click()

def sm_vc_ap_filter_by_status():
    # Filter by status #
    try :
        time.sleep(3)
        
        list_status_before=[]
        if cm.is_Displayed(pr_mn.mn_pro["list_request"]) == True :
            cm.xlsx(pr_mn.mn_filter["pass"] ,pr_mn.msg_mn["pass_no_filter"])
        else:
            # Get all status from list request vacation #
            i=j=1
            time.sleep(3)
            driver.find_element_by_xpath(pr_mn.mn_pro["ic_to_end_page"]).click()
            current_page_text=driver.find_element_by_xpath(pr_mn.mn_pro["page_current"]).text
            current_page=int(current_page_text)
            driver.find_element_by_xpath(pr_mn.mn_pro["ic_to_first_page"]).click()
            
            while i <= current_page:
                if i == current_page :
                    j=1
                    time.sleep(3)
                    total_request=driver.find_elements_by_xpath(pr_mn.mn_pro["list_re_vc"])
                    while j<= len(total_request):
                        status_name=driver.find_element_by_xpath(pr_mn.vc_ap["cl_status1"]+str(j)+pr_mn.vc_ap["cl_status2"]).text
                        list_status_before.append(status_name)
                        j=j+1
                else:
                    while j<= 20:
                        status_name=driver.find_element_by_xpath(pr_mn.vc_ap["cl_status1"]+str(j)+pr_mn.vc_ap["cl_status2"]).text
                        list_status_before.append(status_name)
                        j=j+1
                time.sleep(2)
                total_ic=driver.find_elements_by_xpath(pr_mn.mn_pro["total_ic"])
                driver.find_element_by_xpath(pr_mn.mn_pro["ic_next_page"]+ str(len(total_ic)-1)+pr_mn.mn_pro["ic_next_page1"]).click()
                i=i+1
        
            # < Filter by status is Status > #
            result_request = filter_status("Request","select_st_request",list_status_before)
            
            # < Filter by status is Waiting > #
            result_waiting = filter_status("Waiting","select_st_waiting",list_status_before)
            
            # < Filter by status is Progressing > #
            result_progressing = filter_status("Progressing","select_st_progressing",list_status_before)
            
            # < Filter by status is Completed > #
            result_completed =  filter_status("Completed","select_st_completed",list_status_before)
            
            if result_request == True and result_progressing == True and result_completed == True and result_waiting== True:
                cm.xlsx(pr_mn.mn_filter["pass"] ,pr_mn.msg_mn["pass_filter"])
            else:
                cm.xlsx(pr_mn.mn_filter["fail"] ,pr_mn.msg_mn["fail_filter"])
    except:    
        driver.find_element_by_link_text("Vacation Approve").click()          

def submenu_vacation_adjust():
    
    cm.msg("n", "SUB MENU : MANAGER PROCESSING ")
    if cm.is_Displayed1("textlink","Vacation Adjust") == True:

        cm.msg("n", "I.VACATION ADJUST")
        driver.find_element_by_xpath(data["sm_vacation_adjust"]).click()
        if cm.is_Displayed(pr_mn.mn_pro["ip_search_user"]) == True :
            cm.xlsx(pr_mn.mn_sub_adjust["pass"] ,pr_mn.msg_mn["pass_access_sb_adjust"])

            '''
            cm.msg("n", "Search")
            search_user_and_select_user_from_org()
           
            cm.msg("n", "Grant Vacation")
            sm_vc_ad_grant_vacation()
           
            cm.msg("n", "Adjust Vacation")
            sm_vc_ad_adjust_vacation()
            '''
            cm.msg("n", "Register Usage History")
            sm_vc_ad_register_usage_history()

            '''
            cm.msg("n", "Tab Adjust History")
            driver.find_element_by_xpath(pr_mn.mn_pro["tab_adj_history"]).click()
            if cm.is_Displayed(pr_mn.mn_pro["text_adj_history"]) == True : 
                cm.xlsx(mn_tab_adjust["pass"] ,pr_mn.msg_mn["pass_access_tb_his"])
            else:
                cm.xlsx(mn_tab_adjust["fail"] ,pr_mn.msg_mn["fail_access_tb_his"])


            cm.msg("n", "Tab Vacation Adjust")
            driver.find_element_by_xpath(pr_mn.mn_pro["tab_vc_adjust"]).click()
            if cm.is_Displayed(pr_mn.mn_pro["text_vc_adjust"]) == True : 
                cm.xlsx(mn_tab_adjust["pass"] ,pr_mn.msg_mn["pass_access_tb_adj"])
            else:
                cm.xlsx(mn_tab_adjust["fail"] ,pr_mn.msg_mn["fail_access_tb_adj"])
            '''
        else:
            cm.xlsx(pr_mn.mn_sub_adjust["fail"] ,pr_mn.msg_mn["fail_access_sb_adjust"])
    else:
        cm.xlsx(pr_mn.mn_sub_adjust["pass"] ,pr_mn.msg_mn["pass_access_no_adjust"])
    
def submenu_vacation_approve():
   
   
    

    cm.msg("n", "II.VACATION APPROVE")
    if cm.is_Displayed1("textlink","Vacation Approve") == True:
        driver.find_element_by_link_text("Vacation Approve").click()


        cm.msg("n", "Vacation Approve")
        if cm.is_Displayed(pr_mn.mn_pro["sub_cancel_request"]) == True :
            cm.xlsx(pr_mn.mn_sub_approve["pass"] ,pr_mn.msg_mn["pass_access_sb_approve"])
            
            
            cm.msg("n","View detail request")
            sm_vc_ap_view_detail_request()
            
            
            cm.msg("n","Filter by status")
            sm_vc_ap_filter_by_status()
            

            cm.msg("n","Tab Cancel Request")
            driver.find_element_by_xpath(pr_mn.mn_pro["sub_cancel_request"]).click()
            if cm.is_Displayed(pr_mn.mn_pro["bt_refresh"]) == True : 
                cm.xlsx(pr_mn.mn_tab_approve["pass"] ,pr_mn.msg_mn["pass_access_tb_cancel"])
            else:
                cm.xlsx(pr_mn.mn_tab_approve["fail"] ,pr_mn.msg_mn["fail_access_tb_cancel"])

            
            cm.msg("n","Tab Vacation Approve")
            driver.find_element_by_xpath(pr_mn.mn_pro["sub_approve"]).click()
            if cm.is_Displayed(pr_mn.mn_pro["text_approve"]) == True : 
                cm.xlsx(pr_mn.mn_tab_approve["pass"] ,pr_mn.msg_mn["pass_access_tb_approve"])
            else:
                cm.xlsx(pr_mn.mn_tab_approve["fail"] ,pr_mn.msg_mn["fail_access_tb_approve"])
           
        else:
            cm.xlsx(pr_mn.mn_sub_approve["fail"] ,pr_mn.msg_mn["fail_access_sb_approve"])
    else:
        cm.xlsx(pr_mn.mn_sub_approve["pass"] ,pr_mn.msg_mn["pass_no_sb_approve"])

def submenu_vacation_per_user():
    
    cm.msg("n", "III.VACATION PER USER")
    try :
        if cm.is_Displayed1("textlink","Vacations Per User") == True:
            driver.find_element_by_link_text("Vacations Per User").click()
            
            if cm.is_Displayed(pr_mn.mn_pro["sub_per"]) == True :
                cm.xlsx(pr_mn.mn_sub_per["pass"] ,pr_mn.msg_mn["pass_access_sb_per"])
                select_user = select_user_from_depart("per")
                time.sleep(5)
                driver.find_element_by_css_selector(pr_mn.mn_pro["sort_search"]).click()
                time.sleep(5)
                if select_user == True :
                    print("test")
                   
                    
                else :
                    print("hi")
            else:
                cm.xlsx(pr_mn.mn_sub_per["fail"] ,pr_mn.msg_mn["fail_access_sb_per"])
        else:
            cm.xlsx(pr_mn.mn_sub_per["pass"] ,pr_mn.msg_mn["pass_no_sb_per"])
    except :
        pass

def submenu_settlement_management():
    cm.msg("n", "IV.SETTLEMENT MANAGEMENT")
    if cm.is_Displayed1("textlink","Settlement Management") == True:
        driver.find_element_by_link_text("Settlement Management").click()
       
        if cm.is_Displayed(pr_mn.mn_pro["sub_settlement"]) == True :
            cm.xlsx(pr_mn.mn_sub_sett["pass"] ,pr_mn.msg_mn["pass_access_sb_mana"])

            cm.msg("n","Tab Vacation Request")
            driver.find_element_by_xpath(pr_mn.mn_pro["sub_sett_request"]).click()
            if cm.is_Displayed(pr_mn.mn_pro["text_sett_request"]) == True : 
                cm.xlsx(pr_mn.mn_tab_sett["pass"] ,pr_mn.msg_mn["pass_access_tb_re"])
            else:
                cm.xlsx(pr_mn.mn_tab_sett["fail"] ,pr_mn.msg_mn["fail_access_tb_re"])

            
            cm.msg("n","Change History")
            driver.find_element_by_xpath(pr_mn.mn_pro["sub_sett_history"]).click()
            if cm.is_Displayed(pr_mn.mn_pro["text_sett_history"]) == True : 
                cm.xlsx(pr_mn.mn_tab_sett["pass"] ,pr_mn.msg_mn["pass_access_tb_cha"])
            else:
                cm.xlsx(pr_mn.mn_tab_sett["fail"] ,pr_mn.msg_mn["fail_access_tb_cha"])

            cm.msg("n","Settlement")
            driver.find_element_by_xpath(pr_mn.mn_pro["sub_sett"]).click()
            if cm.is_Displayed(pr_mn.mn_pro["text_sett"]) == True : 
                cm.xlsx(pr_mn.mn_tab_sett["pass"] ,pr_mn.msg_mn["pass_access_tb_set"])
            else:
                cm.xlsx(pr_mn.mn_tab_sett["fail"] ,pr_mn.msg_mn["fail_access_tb_set"])

        else:
            cm.xlsx(pr_mn.mn_sub_sett["fail"] ,pr_mn.msg_mn["fail_access_sb_mana"])
    else:
        cm.xlsx(pr_mn.mn_sub_sett["pass"] ,pr_mn.pr_mn.msg_mn["pass_no_sb_mana"])

def manager_st():
    #submenu_vacation_adjust()
    #submenu_vacation_approve()
    submenu_vacation_per_user()
    #submenu_settlement_management()