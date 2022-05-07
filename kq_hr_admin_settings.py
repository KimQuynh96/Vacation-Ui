from posixpath import abspath
from kq_hr_login import driver,Keys
from kq_hr_lib_common import data,cm,time,WebDriverWait,By,EC,json,Keys
from kq_hr_param import pr_ad
from kq_hr_lib_admin import xpath


def select_user_from_depart():
    # select user from org #
    list_department = driver.find_elements_by_xpath(data["rq_vc"]["list_depart_cc"])    
    total_department = cm.total_data(list_department)
    for i in range(1,total_department):
        time.sleep(1)
        depart_has_user = cm.is_Displayed(xpath.depart_has_user(i)) 
        if depart_has_user == True :
            driver.find_element_by_xpath(xpath.depart_has_user(i)).click()
            total_user =  driver.find_elements_by_xpath(data["rq_vc"]["list_user"])
            for i in range(1,len(total_user)+1):
                is_user = cm.is_Displayed(xpath.is_user(i)) 
                if is_user == True:
                    selected_cc_name = driver.find_element_by_xpath(xpath.cc_name(i)).text
                    driver.find_element_by_xpath(xpath.select_cc(i)).click()
                    return selected_cc_name      
    return False

def filter_type_vacation(list_vc_type_before,select_status,sta_name_to_filter):
    total_type_vacation_to_filter = 0
    result_filter = True
    xpath.open_filter(select_status)
    
    # Count type vacation from all type #
    for i in list_vc_type_before:
        if i  == sta_name_to_filter:
            total_type_vacation_to_filter= total_type_vacation_to_filter + 1

    # Result filter :Filtered list is empty #
    time.sleep(3)
    list_vacation = driver.find_elements_by_xpath(pr_ad.dt_ad["list_vacation"])
    if len(list_vacation) == 0:
        # Total type to filter from list before search = 0 #
        if total_type_vacation_to_filter == 0 :
            cm.msg("p" ,xpath.filter_text("p" , sta_name_to_filter))
        else:
            result_filter = False
            cm.xlsx(pr_ad.filter_status["fail"] ,xpath.filter_text("f" , sta_name_to_filter))

    # Filtered list is not empty  = > Check filter list contains other type  #
    else:
        time.sleep(3)
        j = 1
        rows = driver.find_elements_by_xpath(pr_ad.dt_ad["list_vacation"])
        total_request = cm.total_data(rows)
        while j <= total_request :
            vc_type = driver.find_element_by_xpath(xpath.vc_type(j)).text
            if vc_type != sta_name_to_filter:
                result_filter = False
                cm.xlsx(pr_ad.filter_status["fail"] ,xpath.filter_text("f" , sta_name_to_filter))
                break
            j = j+1
        if result_filter == True:
            cm.msg("p" ,xpath.filter_text("p" , sta_name_to_filter))

    return result_filter

def get_list_manager():
    # get manager list with permission #
    
    all_manager = []
    i = 1
    time.sleep(3)
    list_user_name = driver.find_elements_by_xpath(pr_ad.dt_ad["list_manager"])
    total_manager = cm.total_data(list_user_name)
    total = {"total_manager":total_manager}
    all_manager.append(total)
    while i <= total_manager:
        list_manager = xpath.par_manager()
        list_manager["no"] = i
        list_manager["name"] = driver.find_element_by_xpath(xpath.manager("ma",i)).text
        list_manager["approve"] = driver.find_element_by_xpath(xpath.manager("ap",i)).text
        list_manager["adjust"] = driver.find_element_by_xpath(xpath.manager("ad",i)).text
        list_manager["settlement"] = driver.find_element_by_xpath(xpath.manager("se",i)).text
        all_manager.append(list_manager)
        i = i+1
    return all_manager

def sm_cre_vc_create_vacation():
    # Create Vacation #
    #try:
    param = json.loads(cm.param_data())
    driver.find_element_by_xpath(pr_ad.dt_ad["bt_create"]).click() 
    if cm.is_Displayed(pr_ad.dt_ad["bt_next"]) == True :
        cm.msg("p" ,pr_ad.msg_ad["pass_create_vc_button"])
        
        # Enter vacation name #
        vc_name = driver.find_element_by_xpath(pr_ad.dt_ad["ip_name"])
        vc_name.send_keys(param["vacation_name"])
        time.sleep(1)
        if vc_name.get_attribute('value') == param["vacation_name"] :
            cm.msg("p" ,pr_ad.msg_ad["pass_create_vc_name"])
            result_vc_name = True
        else:
            result_vc_name = False
            cm.xlsx(pr_ad.ad_create["fail"] ,pr_ad.msg_ad["fail_create_vc_name"])

        # Enter number of days vacation #
        driver.find_element_by_xpath(pr_ad.dt_ad["bt_next"]).click()
        if cm.is_Displayed(pr_ad.dt_ad["number_day_off"]) ==True :
            cm.msg("p" ,pr_ad.msg_ad["pass_create_vc_next"])
            driver.implicitly_wait(10)
            day_off = driver.find_element_by_xpath(pr_ad.dt_ad["number_day_off"])
            day_off.clear()
            day_off.send_keys(param["number_day_off"])
            driver.implicitly_wait(10)
            if str(day_off.get_attribute('value')) == str(param["number_day_off"]) :
                cm.msg("p" , pr_ad.msg_ad["pass_create_vc_days"])
            else:
                cm.xlsx(pr_ad.ad_create["fail"] ,pr_ad.msg_ad["fail_create_vc_days"])
            
        else:
            cm.xlsx(pr_ad.ad_create["fail"] ,pr_ad.msg_ad["fail_create_vc_next"])

        # Save vacation #
        if result_vc_name == True :
            driver.find_element_by_xpath(pr_ad.dt_ad["bt_save"]).click()
            if cm.is_Displayed(pr_ad.dt_ad["vacation_list"]) == True : 
                cm.xlsx(pr_ad.ad_create["pass"] ,pr_ad.msg_ad["pass_create_vc"])

                driver.find_element_by_xpath(pr_ad.dt_ad["vacation_list"]).click()
                time.sleep(1)
                driver.find_element_by_xpath(pr_ad.dt_ad["bt_reresh"]).click()
        
                # Check created vacation is in vacation list #
                i = 1
                result_save = False
                time.sleep(3)
                list_vacation = driver.find_elements_by_xpath(pr_ad.dt_ad["list_vacation"])
                total_vacation = cm.total_data(list_vacation)
                if total_vacation == 0:
                    cm.xlsx(pr_ad.ad_create["fail"] ,pr_ad.msg_ad["fail_create_displayed"])
                else:
                    while i <= total_vacation:
                        vc_name = driver.find_element_by_xpath(cm.xpath2(pr_ad.dt_ad["vc_name"],str(i),pr_ad.dt_ad["vc_name1"])).text
                        if vc_name == param["vacation_name"] :
                            cm.xlsx(pr_ad.ad_create["pass"] ,pr_ad.msg_ad["pass_create_displayed"])
                            result_save = True
                            break
                        i = i+1
                    if result_save == False:
                        cm.xlsx(pr_ad.ad_create["fail"],pr_ad.msg_ad["fail_create_displayed"])
                
            else:
                cm.xlsx(pr_ad.ad_create["fail"],pr_ad.msg_ad["fail_create_vc_save"])
    else:
        cm.xlsx(pr_ad.ad_create["fail"],pr_ad.msg_ad["fail_create_vc_click"])
    driver.find_element_by_link_text("Create Vacation").click()
    #except:
        #driver.find_element_by_link_text("Create Vacation").click()
             
def sm_cre_vc_delete_vacation():
    # Delete vacation #
    try:
        i = 1
        result_delete_vc = False
        result_remove_vc = True
        delete_vc = pr_ad.param_excel["delete_vc"]
        time.sleep(3)
        list_vacation = driver.find_elements_by_xpath(pr_ad.dt_ad["list_vacation"])
        total_vc =   cm.total_data(list_vacation)
    
        if total_vc == 0 :
            cm.xlsx(delete_vc["pass"],pr_ad.msg_ad["pass_delete_no_vc"])
        else:
            # Choose vacation to delete #
            while i <= total_vc:
                vc_name_to_delete = driver.find_element_by_xpath(xpath.vc_name_delete(i)).text
                driver.find_element_by_xpath(pr_ad.dt_ad["bt_delete_vc"]).click()
                if cm.is_Displayed(pr_ad.dt_ad["bt_delete_vc2"]) == True :
                    cm.msg("p" ,pr_ad.msg_ad["pass_delete_icon"])
                    driver.find_element_by_xpath(pr_ad.dt_ad["bt_delete_vc2"]).click()
                    if cm.is_Displayed(pr_ad.dt_ad["bt_close"]) == False :
                        cm.xlsx(delete_vc["pass"] ,pr_ad.msg_ad["pass_delete_vc"])
                        result_delete_vc= True
                        break
                    else:
                        cm.xlsx(delete_vc["fail"] ,pr_ad.msg_ad["fail_delete_vc"])

                else:
                    cm.xlsx(delete_vc["fail"] ,pr_ad.msg_ad["fail_delete_icon"])
                i=i+1

            # If the deletion is successful, check the vacation that has been deleted from the vacation list #
            if result_delete_vc == True:
                time.sleep(3)
                list_vacation1 = driver.find_elements_by_xpath(pr_ad.dt_ad["list_vacation"])
                total_vc1 = cm.total_data(list_vacation1)

                # total vacation before deletion = 0 #
                if total_vc1 == 0 :
                    # total vacation before deletion #
                    if total_vc ==1 :
                        cm.xlsx(delete_vc["pass"] ,pr_ad.msg_ad["pass_delete_removed"])
                    else :
                        cm.xlsx(delete_vc["fail"] ,pr_ad.msg_ad["fail_delete_removed"])
                else:
                    while i <= total_vc1:
                        vc_name_1 = driver.find_element_by_xpath(xpath.vc_name(i)).text
                        if vc_name_to_delete == vc_name_1:
                            cm.xlsx(delete_vc["fail"] ,pr_ad.msg_ad["fail_delete_removed"])
                            result_remove_vc = False
                            break
                        i = i+1
                    if result_remove_vc == True:
                        cm.xlsx(delete_vc["pass"] ,pr_ad.msg_ad["pass_delete_removed"])
    except:
        driver.find_element_by_link_text("Create Vacation").click()
            
def sm_cre_vc_filter_vacation():
    # Filter Vacation #
    try:
        list_vc_type_before = []
        i = 1
        # Get all vacation type from list vacation #
        time.sleep(3)
        list_vacation2 =   driver.find_elements_by_xpath(pr_ad.dt_ad["list_vacation"])
        number_vc1 =   len(list_vacation2)

        if number_vc1 == 0:
            cm.xlsx(pr_ad.filter_status["pass"] ,pr_ad.msg_ad["pass_filter_no_re"])
        else:
            while i <= number_vc1:
                vc_type = driver.find_element_by_xpath(cm.xpath2(pr_ad.dt_ad["vc_type"],str(i),pr_ad.dt_ad["vc_type1"])).text
                list_vc_type_before.append(vc_type)
                i = i+1

            # Filter by vacation type is Regular #
            resullt_regular = filter_type_vacation(list_vc_type_before,"type_regular","Regular Vacation")
            
            # Filter by vacation type is Grant #
            resullt_grant = filter_type_vacation(list_vc_type_before,"type_grant","Grant Vacation")

            # Filter by vacation type is Other #
            resullt_other = filter_type_vacation(list_vc_type_before,"type_other","Other Vacation")

            if resullt_regular == True and resullt_grant == True and resullt_other == True :
                cm.xlsx(pr_ad.filter_status["pass"] ,pr_ad.msg_ad["pass_filter_vc"])
            else:
                cm.xlsx(pr_ad.filter_status["fail"] ,pr_ad.msg_ad["fail_filter_vc"])

        driver.find_element_by_link_text("Create Vacation").click()
    except:
        driver.find_element_by_link_text("Create Vacation").click()

def sm_cre_vc_view_detail_vacation():

    # View Detail Vacation #
    try :
        time.sleep(3)
        list_vacation = driver.find_elements_by_xpath(pr_ad.dt_ad["list_vacation"])
        total_vc = len(list_vacation)

        if total_vc <1 :
            cm.xlsx(pr_ad.ad_detail["pass"] ,pr_ad.msg_ad["pass_view_vc"])
        else:
            driver.find_element_by_xpath(pr_ad.dt_ad["bt_detail_vc1"]).click()
            if cm.is_Displayed(pr_ad.dt_ad["ra_use"]) == True :
                cm.msg("p" ,pr_ad.msg_ad["pass_view_icon"])


                cm.msg("n" ,"Tab Detail Settings")
                if cm.is_Displayed(pr_ad.dt_ad["tab_det_st"]) == True :
                    driver.find_element_by_xpath(pr_ad.dt_ad["tab_det_st"]).click()
                    if cm.is_Displayed(pr_ad.dt_ad["check_det"]) == True :
                        cm.xlsx(pr_ad.ad_tab_detail["pass"] ,pr_ad.msg_ad["pass_access_tb_detail"])
                    else:
                        cm.xlsx(pr_ad.ad_tab_detail["fail"] ,pr_ad.msg_ad["fail_access_tb_detail"])
                else:
                    cm.msg("p",pr_ad.msg_ad["pass_no_tb_detail"])


                cm.msg("n" ,"Tab Default Settings")
                if cm.is_Displayed(pr_ad.dt_ad["tab_def_st"]) == True :
                    driver.find_element_by_xpath(pr_ad.dt_ad["tab_def_st"]).click()
                    if cm.is_Displayed(pr_ad.dt_ad["check_def"]) == True :
                        cm.xlsx(pr_ad.ad_tab_detail["pass"] ,pr_ad.msg_ad["pass_access_tb_def"])
                    else:
                        cm.xlsx(pr_ad.ad_tab_detail["fail"] ,pr_ad.msg_ad["fail_access_tb_def"])
                else:
                    cm.msg("p" ,pr_ad.msg_ad["pass_no_tb_def"])
                driver.find_element_by_link_text("Create Vacation").click()
            
            else:
                cm.xlsx(pr_ad.ad_detail["fail"] ,pr_ad.msg_ad["fail_detail_icon"])

            driver.find_element_by_link_text("Create Vacation").click()
    except:
        driver.find_element_by_link_text("Create Vacation").click()

def sm_cre_vc_search_vacation():
    # Search vacation #
    try:
        j = 1
        result_search = True
       
        # Get vacation list #
        time.sleep(3)
        list_vacation = driver.find_elements_by_xpath(pr_ad.dt_ad["list_vacation"])
        total_vacation = cm.total_data(list_vacation)

        if total_vacation < 1:
            cm.xlsx(pr_ad.ad_search["pass"] ,pr_ad.msg_ad["pass_search_no"])
        else :
            vacation_to_search = driver.find_element_by_xpath(pr_ad.dt_ad["search_vc3"]).text
            # Enter vacation name to search #
            input_search = driver.find_element_by_xpath(pr_ad.dt_ad["ip_search_vc"])
            input_search.send_keys(vacation_to_search)
            input_search.send_keys(Keys.RETURN)

            # Check result search #
            time.sleep(3)
            list_vc_after_search = driver.find_elements_by_xpath(pr_ad.dt_ad["list_vacation"])
            total_va_after_search = len(list_vc_after_search)

            # total vacation after search = 0 #
            if total_va_after_search == 0 :
                cm.xlsx(pr_ad.ad_search["fail"] ,pr_ad.msg_ad["fail_search_vc"])
            else:
                # Check the search list results that contain the search keyword #
                while j <= total_va_after_search:
                    vacation_name1 = driver.find_element_by_xpath(xpath.search()).text
                    if vacation_to_search != vacation_name1 :
                        result_search = False
                        cm.xlsx(pr_ad.ad_search["fail"] ,pr_ad.msg_ad["fail_search_vc"])
                        break
                    j = j+1
                if result_search == True :
                        cm.xlsx(pr_ad.ad_search["pass"],pr_ad.msg_ad["pass_search_vc"])

        driver.find_element_by_link_text("Create Vacation").click()
    except:
        driver.find_element_by_link_text("Create Vacation").click()
def sm_mn_se_add_manager():
    # Add Manager #
    try :
        manager_name = "TS2"
       

        driver.find_element_by_xpath(pr_ad.dt_ad["bt_add_manager"]).click()
        if cm.is_Displayed(pr_ad.dt_ad["ip_search_user"]) ==True :
            cm.msg("p" ,pr_ad.msg_ad["pass_manager_button"])

            # Select user from search #
            ip_search_user = driver.find_element_by_xpath(pr_ad.dt_ad["ip_search_user"])
            ip_search_user.send_keys(manager_name)
            ip_search_user.send_keys(Keys.RETURN)
            if ip_search_user.get_attribute('value') == manager_name :
                cm.msg("p" ,pr_ad.msg_ad["pass_manager_enter"])
                time.sleep(2)
                list_search_mn = driver.find_elements_by_xpath(pr_ad.dt_ad["list_search_mn"])
                if len(list_search_mn) != 0 :
                    driver.find_element_by_css_selector(pr_ad.dt_ad["select_user"]).click()
                    cm.msg("p" ,pr_ad.msg_ad["pass_manager_select"])

                    driver.find_element_by_xpath(pr_ad.dt_ad["bt_add_user"]).click()
                    list_mn = driver.find_elements_by_xpath(pr_ad.dt_ad["mn_selected"])
                    total_mn = cm.total_data(list_mn)

                    if total_mn == 0 :
                        cm.xlsx(pr_ad.ad_manager["fail"] ,pr_ad.msg_ad["fail_manager_add"])
                    else:
                        mn_name = driver.find_element_by_xpath(pr_ad.dt_ad["mn_name"]).text
                        if mn_name == manager_name:
                            cm.msg("p" ,pr_ad.msg_ad["pass_manager_add"])

                            driver.find_element_by_xpath(pr_ad.dt_ad["bt_save_user"]).click()
                            if cm.is_Displayed(pr_ad.dt_ad["bt_add_manager"]) ==True :
                                cm.xlsx(pr_ad.ad_manager["pass"] ,pr_ad.msg_ad["pass_manager"])
                                
                                # Add manager successfully , check the manager displayed in the manager list #
                                time.sleep(3)
                                list_user_name = driver.find_elements_by_xpath(pr_ad.dt_ad["list_manager"])
                                total_manager = cm.total_data(list_user_name)
                                check_manager = False
                                i = 1

                                # Total vacation after added = 0 #
                                if total_manager == 0 :
                                    cm.xlsx(pr_ad.ad_manager["fail"] ,pr_ad.msg_ad["fail_manager_displayed"])

                                #check added vacation is in vacation list #
                                else:
                                    while i <= total_manager:
                                        name =driver.find_element_by_xpath(xpath.manager_name(i)).text
                                        if name.strip() == manager_name.strip():
                                            cm.xlsx(pr_ad.ad_manager["pass"] ,pr_ad.msg_ad["pass_manager_displayed"])
                                            check_manager = True
                                            break
                                        i = i+1
                                    if check_manager == False:
                                        cm.xlsx(pr_ad.ad_manager["fail"] ,pr_ad.msg_ad["fail_manager_displayed"])
                            else:
                                cm.xlsx(pr_ad.ad_manager["fail"] ,pr_ad.msg_ad["fail_manager"])
                else:
                    cm.xlsx(pr_ad.ad_manager["pass"] ,pr_ad.msg_ad["pass_search_no_exist"])
            else:
                cm.xlsx(pr_ad.ad_manager["fail"] ,pr_ad.msg_ad["fail_manager_enter"])
        else :
            cm.xlsx(pr_ad.ad_manager["fail"] ,pr_ad.msg_ad["fail_manager_button"])
    except:
        driver.find_element_by_xpath(data["sm_manager_settings"]).click()

def sm_mn_se_search_manager():
    # Search Manager #
    j = 1
    result_search = True

    time.sleep(3)
    list_manager = driver.find_elements_by_xpath(pr_ad.dt_ad["list_manager"])
    total_manager = cm.total_data(list_manager)
    if total_manager < 1:
        cm.xlsx(pr_ad.ad_search_mn["pass"] ,pr_ad.msg_ad["pass_search_no_mn"])
    else:
        manager_to_search = driver.find_element_by_xpath(pr_ad.dt_ad["mn_search"]).text.strip()

        # Enter manager name to search #
        input_search = driver.find_element_by_xpath(pr_ad.dt_ad["ip_search_vc"])
        input_search.send_keys(manager_to_search)
        input_search.send_keys(Keys.RETURN)

        # Check result search #
        time.sleep(3)
        list_vc_after_search = driver.find_elements_by_xpath(pr_ad.dt_ad["list_manager"])
        total_mn_after_search = len(list_vc_after_search)

        # After search , the search result = 0 #
        if total_mn_after_search == 0 :
            cm.xlsx(pr_ad.ad_search_mn["fail"] ,pr_ad.msg_ad["fail_search_mn"])
        
        # Check manager is in search list #
        else:
            while j <= total_mn_after_search:
                manager_name = driver.find_element_by_xpath(cm.xpath("tr",str(j),pr_ad.dt_ad["manager"])).text.strip()
                if manager_name != manager_to_search :
                    result_search = False
                    cm.xlsx(pr_ad.ad_search_mn["fail"] ,pr_ad.msg_ad["fail_search_mn"])
                    break
                j = j+1
            if result_search == True :
                    cm.xlsx(pr_ad.ad_search_mn["pass"],pr_ad.msg_ad["pass_search_mn"])

def sm_mn_se_modify_permission():
    # Modify Manager'permission #
    try:
        
        all_manager_before = get_list_manager()

        if all_manager_before[0]["total_manager"] == 0:
            cm.xlsx(pr_ad.ad_modify_mn["pass"] ,pr_ad.msg_ad["pass_modify_no_mn"])

        else:
            manager_before_modify = all_manager_before[1]
            manager_to_modify = manager_before_modify["name"]

            time.sleep(3)
            driver.find_element_by_css_selector(xpath.ic_edit(manager_before_modify)).click()
            if cm.is_Displayed(pr_ad.dt_ad["bt_save_edit"]) == True :
                cm.msg("p" ,pr_ad.msg_ad["pass_modify_icon"])

                time.sleep(3)
                driver.find_element_by_xpath(pr_ad.dt_ad["cb_adjust_authority"]).click()
                cm.msg("p" ,pr_ad.msg_ad["pass_modify_adjust"])
                time.sleep(1)

                driver.find_element_by_xpath(pr_ad.dt_ad["cb_approve_authority"]).click()
                cm.msg("p" ,pr_ad.msg_ad["pass_modify_approve"])
                time.sleep(1)

                driver.find_element_by_xpath(pr_ad.dt_ad["cb_settlement_authority"]).click()
                cm.msg("p" ,pr_ad.msg_ad["pass_modify_settlement"])
                time.sleep(1)

                driver.find_element_by_xpath(pr_ad.dt_ad["bt_save_edit"]).click()
                if cm.is_Displayed(pr_ad.dt_ad["bt_refresh"]) ==True : 
                    cm.xlsx(pr_ad.ad_modify_mn["pass"] ,pr_ad.msg_ad["pass_modify_permission"])

                    # Get manager list #
                    all_manager_after = get_list_manager()
                    manager_after_modify = ""
                    for no,manager in enumerate(all_manager_after):
                        if no > 0 :
                            manager["name"] == manager_to_modify
                            manager_after_modify = manager
                            break
                    # Find the edited manager to check edited permissions #
                    if manager_after_modify != "" :
                        result_approve = False
                        result_adjust = False
                        result_settlement = False
                        if manager_after_modify["approve"] != manager_before_modify["approve"]:
                            cm.msg("p" ,pr_ad.msg_ad["pass_approve_modifyed"])
                            result_approve = True
                        else:
                            cm.xlsx(pr_ad.ad_modify_mn["fail"] ,pr_ad.msg_ad["fail_approve_modifyed"])

                        if manager_after_modify["adjust"] != manager_before_modify["adjust"]:
                            result_adjust = True
                            cm.msg("p" ,pr_ad.msg_ad["pass_adjust_modifyed"])
                        else:
                            cm.xlsx(pr_ad.ad_modify_mn["fail"] ,pr_ad.msg_ad["fail_adjust_modifyed"])

                        if manager_after_modify["settlement"] != manager_before_modify["settlement"]:
                            result_settlement = True
                            cm.msg("p" ,pr_ad.msg_ad["pass_settlemen_modifyed"])
                        else:
                            cm.xlsx(pr_ad.ad_modify_mn["fail"] ,pr_ad.msg_ad["fail_settlemen_modifyed"])

                        if result_approve == True and result_adjust == True and result_settlement == True :
                            cm.xlsx(pr_ad.ad_modify_mn["pass"] ,pr_ad.msg_ad["pass_modify"])

                        else:
                            cm.xlsx(pr_ad.ad_modify_mn["fail"] ,pr_ad.msg_ad["fail_modify"])
                    else:
                        cm.xlsx(pr_ad.ad_modify_mn["fail"] ,pr_ad.msg_ad["fail_modify_no_find"])
                else:
                    cm.xlsx(pr_ad.ad_modify_mn["fail"] ,pr_ad.msg_ad["fail_modify_save"] )
            else:
                cm.xlsx(pr_ad.ad_modify_mn["fail"] ,pr_ad.msg_ad["fail_modify_icon"])
    except:
        driver.find_element_by_xpath(data["sm_manager_settings"]).click()

def sm_mn_se_delete_manager():
    # Delete manager #
    try:
        result_delete = True
       
        all_manager_before  = get_list_manager()

        if all_manager_before[0]["total_manager"] == 0:
            cm.xlsx(pr_ad.ad_delete_mn["pass"] ,pr_ad.msg_ad["pass_delete_mn_no"])
        else:
            manager_before_delete = all_manager_before[1]
            manager_to_delete = manager_before_delete["name"]
            driver.find_element_by_css_selector(xpath.ic_delete(manager_before_delete)).click()
            if cm.is_Displayed(pr_ad.dt_ad["bt_delete"]) == True : 
                cm.msg("p" ,pr_ad.msg_ad["pass_delete_mn_icon"])

                driver.find_element_by_xpath(pr_ad.dt_ad["bt_delete"]).click()
                if cm.is_Displayed(pr_ad.dt_ad["bt_refresh"]) ==True : 
                    cm.xlsx(pr_ad.ad_delete_mn["pass"] ,pr_ad.msg_ad["pass_delete_mn"])

                    # Delete manager successfully , check the manager removed from the manager list #
                    all_manager_after = get_list_manager()

                    # total manager after deleted #
                    if cm.total_data(all_manager_after) == 0 :

                        # total manager before delete #
                        if all_manager_before[0]["total_manager"] == 1:
                            cm.xlsx(pr_ad.ad_delete_mn["pass"] ,pr_ad.msg_ad["pass_delete_mn_removed"])
                        else :
                            cm.xlsx(pr_ad.ad_delete_mn["fail"] ,pr_ad.msg_ad["fail_delete_mn_removed"])

                    else:
                        for no,manager in enumerate(all_manager_after):
                            if no > 0 and manager["name"] == manager_to_delete :
                                result_delete = False
                                break
                        if result_delete == True:
                            cm.xlsx(pr_ad.ad_delete_mn["pass"] ,pr_ad.msg_ad["pass_delete_mn_removed"])
                        else:
                            cm.xlsx(pr_ad.ad_delete_mn["fail"] ,pr_ad.msg_ad["fail_delete_mn_removed"] )
                else:
                    cm.xlsx(pr_ad.ad_delete_mn["fail"] ,pr_ad.msg_ad["fail_delete_mn"]) 
            else:
                cm.xlsx(pr_ad.ad_delete_mn["fail"] ,pr_ad.msg_ad["fail_delete_mn_icon"]) 
    except:
        driver.find_element_by_xpath(data["sm_manager_settings"]).click()      

def sm_mn_se_add_arbitrary_decision():
    # Add user is Arbitrary Decision #
    try:
        time.sleep(2)
        arbitrary_decision ="TS3"
        driver.find_element_by_xpath(pr_ad.dt_ad["bt_select_approver"]).click()
        if cm.is_Displayed(pr_ad.dt_ad["ip_search_user"]) == True : 
            cm.msg("p" ,pr_ad.msg_ad["pass_Approver_button"])

            # Select user from search to add#
            ip_search_user = driver.find_element_by_xpath(pr_ad.dt_ad["ip_search_user"])
            ip_search_user.send_keys("TS3")
            ip_search_user.send_keys(Keys.RETURN)
            if ip_search_user.get_attribute('value') == arbitrary_decision:
                cm.msg("p" ,pr_ad.msg_ad["pass_approver_enter"])
                
                time.sleep(2)
                list_search_mn=driver.find_elements_by_xpath(pr_ad.dt_ad["list_search_mn"])
                if len(list_search_mn) != 0 :
                    time.sleep(1)
                    driver.find_element_by_css_selector(pr_ad.dt_ad["select_user"]).click()
                    cm.msg("p" ,pr_ad.msg_ad["pass_approver_select"])

                
                    driver.find_element_by_xpath(pr_ad.dt_ad["bt_add_arbitrary"]).click()
                    cm.msg("p",pr_ad.msg_ad["pass_approver_org"])

                    driver.find_element_by_xpath(pr_ad.dt_ad["bt_save_user"]).click()
                    if cm.is_Displayed(pr_ad.dt_ad["bt_select_approver"]) == True :
                        cm.xlsx(pr_ad.ad_arbitrary["pass"] ,pr_ad.msg_ad["pass_approver_add"])

                        # Check Added Arbitrary Decision #
                        time.sleep(3)
                        list_ardec = driver.find_elements_by_xpath(pr_ad.dt_ad["list_ardec1"])
                        total_ardec = cm.total_data(list_ardec)
                        check_ardec = False

                        # Total Arbitrary Decision after add #
                        if total_ardec == 0 :
                            cm.xlsx(pr_ad.ad_arbitrary["pass"] ,pr_ad.msg_ad["pass_approver_displayed"])
                        else:
                            # check Arbitrary Decision is in the list #
                            i = 1
                            while i <= total_ardec:
                                name_ardec = driver.find_element_by_xpath(xpath.ardec_name(i)).text
                                if name_ardec.strip() == arbitrary_decision.strip():
                                    cm.xlsx(pr_ad.ad_arbitrary["pass"] ,pr_ad.msg_ad["pass_approver_displayed"])
                                    check_ardec = True
                                    break
                                i = i+1

                            if check_ardec == False:
                                cm.xlsx(pr_ad.ad_arbitrary["fail"] ,pr_ad.msg_ad["fail_approver_displayed"])
                    else:
                        cm.xlsx(pr_ad.ad_arbitrary["fail"] ,pr_ad.msg_ad["fail_approver_add"])
            else:
                cm.xlsx(pr_ad.ad_arbitrary["fail"] ,pr_ad.msg_ad["fail_approver_enter"])
        else:
            cm.xlsx(pr_ad.ad_arbitrary["fail"] ,pr_ad.msg_ad["fail_approver_button"])
    except:
        driver.find_element_by_xpath(data["sm_manager_settings"]).click()
        driver.find_element_by_xpath(pr_ad.dt_ad["mn_approval_settings"]).click()

def sm_mn_se_delete_arbitrary_decision():
    # Delete user is Arbitrary Decision -#
    
    result_delete = False
    removed = True


    time.sleep(3)
    list_ardec = driver.find_elements_by_xpath(pr_ad.dt_ad["list_ardec1"])
    total_ardec_before = cm.total_data(list_ardec)
    if total_ardec_before  == 0 :
        cm.xlsx(pr_ad.ad_delete_arbitrary["pass"] ,pr_ad.msg_ad["pass_delete_no_ap"])
    else:
        i = 1
        while i <= total_ardec_before:
            ardec_to_delete = driver.find_element_by_xpath(xpath.delete_ardec(i)).text
            driver.find_element_by_css_selector(pr_ad.dt_ad["ic_dl_ardec"]).click()

            if cm.is_Displayed(pr_ad.dt_ad["bt_close_ar"]) == True :
                cm.msg("p" ,pr_ad.msg_ad["pass_delete_ap_click"])
                driver.find_element_by_xpath(pr_ad.dt_ad["bt_dele_ar"]).click()

                if cm.is_Displayed(pr_ad.dt_ad["bt_close_ar"]) == False :
                    cm.xlsx(pr_ad.ad_delete_arbitrary["pass"] ,pr_ad.msg_ad['pass_delete_ap'])
                    result_delete = True
                    break
                else:
                    cm.xlsx(pr_ad.ad_delete_arbitrary["fail"] ,pr_ad.msg_ad["fail_delete_ap"])
                break
            else:
                cm.msg("f",pr_ad.msg_ad["fail_delete_ap_click"])
                
            i = i+1


        if result_delete == True :
            time.sleep(3)
            list_ardec = driver.find_elements_by_xpath(pr_ad.dt_ad["list_ardec1"])
            total_ardec_after = cm.total_data(list_ardec)
            i = 1
            # total Arbitrary Decision after delete = 0 #
            if total_ardec_after == 0 :

                # total vacation before deletion = 1 #
                if total_ardec_before == 1:
                    cm.xlsx(pr_ad.ad_delete_arbitrary["pass"] , pr_ad.msg_ad["pass_delete_ap_removed"])
                else:
                    cm.xlsx(pr_ad.ad_delete_arbitrary["fail"] , pr_ad.msg_ad["fail_delete_ap_removed"])

            else:
                # Check Arbitrary Decision is in the list #
                while i <= total_ardec_after:
                    ardec_name = driver.find_element_by_xpath(xpath.ardec_name(i)).text
                    if ardec_name == ardec_to_delete :
                        removed = False
                        cm.xlsx(pr_ad.ad_delete_arbitrary["fail"] , pr_ad.msg_ad["fail_delete_ap_removed"])
                        break
                    i = i+1
                if removed == True :
                    cm.xlsx(pr_ad.ad_delete_arbitrary["pass"] , pr_ad.msg_ad["pass_delete_ap_removed"])
       
def sm_bs_use_settlement():
       
    try :
        
        # get the currently used status of settlement #
        ip_use_settlement = driver.find_element_by_xpath(pr_ad.dt_ad["ip_use_settlement"])
        time.sleep(3)
        use_settlement_before = ip_use_settlement.is_selected()

        # get the  status of settlement after click use/not use #
        driver.find_element_by_xpath(pr_ad.dt_ad["use_settlement"]).click()
        time.sleep(3)
        use_settlement_after = ip_use_settlement.is_selected()

        # compare status before click and after click #
        if use_settlement_before != use_settlement_after:
            cm.xlsx(pr_ad.ad_use_settlement["pass"] ,pr_ad.msg_ad["pass_use_sett"])
            driver.refresh()

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,data["iframe_vc"])))
            driver.switch_to.frame(driver.find_element_by_xpath(data["iframe_vc"]))
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,data["menu_vc"]))).click()

            # if use settlement #
            if use_settlement_after == True :

                # Check the updated Settlement submenu to UI #
                if cm.is_Displayed1("textlink","Request Settlement") == True:
                    cm.xlsx(pr_ad.ad_use_settlement["pass"] ,pr_ad.msg_ad["pass_updated_ui"])
                else:
                    cm.xlsx(pr_ad.ad_use_settlement["fail"] ,pr_ad.msg_ad["fail_updated_ui"])

            # if not use Settlement #
            else:
                # Check the  Settlement submenu removed from UI #
                if cm.is_Displayed1("textlink","Request Settlement") == False:
                    cm.xlsx(pr_ad.ad_use_settlement["pass"] ,pr_ad.msg_ad["pass_updated_ui"])
                else:
                    cm.xlsx(pr_ad.ad_use_settlement["fail"] ,pr_ad.msg_ad["fail_updated_ui"])
    except:
        driver.find_element_by_xpath(data["sm_basic_settings"]).click()

def sm_bs_use_hour_unit():    
    #[Use Hour unit]#
    try :
        # get the currently used status of hour unit  #
        time.sleep(3)
        ip_use_hour_unit = driver.find_element_by_xpath(pr_ad.dt_ad["ip_use_hour_unit"])
        time.sleep(3)
        use_hour_unit_before = ip_use_hour_unit.is_selected()

        # get the  status of hour unit after click use/not use #
        driver.find_element_by_xpath(pr_ad.dt_ad["use_hour_unit"]).click()
        time.sleep(3)
        use_hour_unit_after = ip_use_hour_unit.is_selected()

        # compare status before click and after click #
        if use_hour_unit_before != use_hour_unit_after:
            cm.xlsx(pr_ad.ad_use_hour["pass"] ,pr_ad.msg_ad["pass_use_unit"])
            driver.refresh()

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,data["iframe_vc"])))
            driver.switch_to.frame(driver.find_element_by_xpath(data["iframe_vc"]))
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,data["menu_vc"]))).click()
            time.sleep(3)

            # Go to UI of create vacation to check UI of hour unit  #
            driver.find_element_by_link_text("Create Vacation").click()
            driver.find_element_by_xpath(pr_ad.dt_ad["bt_create"]).click() 
            cm.scroll()
            time.sleep(3)
            ui_hour_unit = driver.find_elements_by_xpath(pr_ad.dt_ad["check_hour_unit"])

            # if use hour unit #
            if use_hour_unit_after == True :
                # At UI of create vacation will up date UI of hour unit #
                if len(ui_hour_unit) == 2:
                    cm.xlsx(pr_ad.ad_use_hour["pass"] ,pr_ad.msg_ad["pass_updated_ui"])
                else:
                    cm.xlsx(pr_ad.ad_use_hour["fail"] ,pr_ad.msg_ad["fail_updated_ui"])

            # if not use hour unit #
            else:
                # At UI of create vacation will no up date UI of hour unit #
                if len(ui_hour_unit) == 1:
                    cm.xlsx(pr_ad.ad_use_hour["pass"] ,pr_ad.msg_ad["pass_updated_ui"])
                else:
                    cm.xlsx(pr_ad.ad_use_hour["fail"] ,pr_ad.msg_ad["fail_updated_ui"])
            
        else:
            cm.xlsx(pr_ad.ad_use_hour["fail"] ,pr_ad.msg_ad['fail_use_unit'])
    except:
        driver.find_element_by_xpath(data["sm_basic_settings"]).click()
        
def sm_bs_use_vacation_schedule():
    #[Use vacation schedule]#
    try:
        # get the currently used status ofvacation schedule  #
        ip_use_vc_schedule = driver.find_element_by_xpath(pr_ad.dt_ad["ip_use_vc_schedule"])
        time.sleep(3)
        use_vc_schedule_before = ip_use_vc_schedule.is_selected()

        # get the  status of hour unit after click use/not use #
        driver.find_element_by_xpath(pr_ad.dt_ad["use_vc_schedule"]).click()
        time.sleep(3)
        use_vc_schedule_after = ip_use_vc_schedule.is_selected()
        
        # compare status before click and after click #
        if use_vc_schedule_before != use_vc_schedule_after:
            cm.xlsx(pr_ad.ad_use_schedule["pass"] ,pr_ad.msg_ad["pass_use_schedule"])
            driver.refresh()
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,data["iframe_vc"])))
            driver.switch_to.frame(driver.find_element_by_xpath(data["iframe_vc"]))
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,data["menu_vc"]))).click()

            time.sleep(3)
            driver.find_element_by_link_text("Vacation Schedule").click()
            
            # if use vacation schedule #
            if use_vc_schedule_after == True :
                # Check the updated "Vacation Schedule" submenu to UI #
                if cm.is_Displayed(pr_ad.dt_ad["bt_depart"]) == True :
                    cm.xlsx(pr_ad.ad_use_schedule["pass"] ,pr_ad.msg_ad["pass_updated_ui"])
                else:
                    cm.xlsx(pr_ad.ad_use_schedule["fail"] ,pr_ad.msg_ad["fail_updated_ui"])

            # if not use vacation schedule #
            else:
                # Check the not updated "Vacation Schedule" submenu to UI #
                if cm.is_Displayed(pr_ad.dt_ad["bt_depart"]) == False :
                    cm.xlsx(pr_ad.ad_use_schedule["pass"] ,pr_ad.msg_ad["pass_updated_ui"])
                else:
                    cm.xlsx(pr_ad.ad_use_schedule["fail"] ,pr_ad.msg_ad["fail_updated_ui"])

        else:
            cm.xlsx(pr_ad.ad_use_schedule["fail"] ,pr_ad.msg_ad["fail_use_schedule"])
    except:
        driver.find_element_by_xpath(data["sm_basic_settings"]).click()
        
def sm_bs_use_hour_unit_limit():
    try:
        
        ip_use_hu_limit = driver.find_element_by_xpath(pr_ad.dt_ad["ip_use_hu_limit"])
        time.sleep(3)
        use_hu_limit_before = ip_use_hu_limit.is_selected()
        
        driver.find_element_by_xpath(pr_ad.dt_ad["use_hu_limit"]).click()
        time.sleep(3)
        use_hu_limit_after = ip_use_hu_limit.is_selected()
        
        if use_hu_limit_before != use_hu_limit_after:
            cm.xlsx(pr_ad.ad_use_hour_unit["pass"] ,pr_ad.msg_ad["pass_use__limit"])
            driver.refresh()
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,data["iframe_vc"])))
            driver.switch_to.frame(driver.find_element_by_xpath(data["iframe_vc"]))
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,data["menu_vc"]))).click()
            time.sleep(3)
            driver.find_element_by_link_text("Request Vacation").click()
            
        else:
            cm.xlsx(pr_ad.ad_use_hour_unit["fail"] ,pr_ad.msg_ad['fail_use__limit'])
    except:
        driver.find_element_by_xpath(data["sm_basic_settings"]).click()
    
def sm_bs_use_approval_exception():
    #[Approval Exception]#
    try:
        check_apex = False
        result_search = False
        
        approval_exception = data["user"]
        ip_use_ex_approval = driver.find_element_by_xpath(pr_ad.dt_ad["ip_use_ap_exception"])

        time.sleep(2)
        use_ex_approval_before = ip_use_ex_approval.is_selected()

        if use_ex_approval_before == False:
            driver.find_element_by_xpath(pr_ad.dt_ad["use_ap_exception"]).click()
        else :
            
            driver.find_element_by_xpath(pr_ad.dt_ad["bt_add_approval_exception"]).click()
            if cm.is_Displayed(pr_ad.dt_ad["ip_search_user"]) ==True :
                cm.msg("p",pr_ad.msg_ad["pass_exception_add"])
                
                # Search user > Select user from search to add #
                ip_search_user = driver.find_element_by_xpath(pr_ad.dt_ad["ip_search_bs"])
                ip_search_user.click()
                ip_search_user.send_keys(approval_exception)
                if ip_search_user.get_attribute('value') == approval_exception:
                    cm.msg("p",pr_ad.msg_ad["pass_exception_enter"])
                    ip_search_user.send_keys(Keys.RETURN)
                    time.sleep(2)
                    list_search_mn = driver.find_elements_by_xpath(pr_ad.dt_ad["list_search_mn"])
                    if len(list_search_mn) != 0 :
                        driver.find_element_by_css_selector(pr_ad.dt_ad["select_user"]).click()
                        cm.msg("p",pr_ad.msg_ad["pass_exception_select"])
                        result_search = True
                        time.sleep(3)
                    else:
                        cm.xlsx(pr_ad.ad_basic["pass"] ,pr_ad.msg_ad["pass_exception_no_exist"])
                else:
                    cm.xlsx(pr_ad.ad_basic["fail"] ,pr_ad.msg_ad["fail_exception_enter"])
                

                # if the search fails => Select user from org #
                if result_search == False :
                    select_user_org=select_user_from_depart()


                if result_search == True or select_user_org != False :
                    driver.find_element_by_xpath(pr_ad.dt_ad["bt_add_user"]).click()
                    cm.msg("p",pr_ad.msg_ad["fail_exception_org"])

                    driver.find_element_by_xpath(pr_ad.dt_ad["bt_save_user"]).click()
                    if cm.is_Displayed(pr_ad.dt_ad["bt_add_approval_exception"]) == True :
                        cm.xlsx(pr_ad.ad_basic["pass"] ,pr_ad.msg_ad["pass_exception"])
                    
                        
                        time.sleep(2)
                        i = 1
                        list_apex = driver.find_elements_by_xpath(pr_ad.dt_ad["list_apex"])
                        total_apex = cm.total_data(list_apex)

                        # total approval exception after added #
                        if total_apex == 0 :
                            cm.xlsx(pr_ad.ad_basic["fail"] ,pr_ad.msg_ad["fail_exception_displayed"])

                        # Check added Approval exception is in the list #
                        else :
                            while i <= total_apex:
                                name_ax = driver.find_element_by_xpath(xpath.ax_name(i)).text
                                if name_ax.strip() == approval_exception.strip():
                                    cm.xlsx(pr_ad.ad_basic["pass"] ,pr_ad.msg_ad["pass_exception_displayed"])
                                    check_apex = True
                                    break
                                i = i+1
                            if check_apex == False:
                                cm.xlsx(pr_ad.ad_basic["fail"] ,pr_ad.msg_ad["fail_exception_displayed"])
                    else:
                        cm.xlsx(pr_ad.ad_basic["fail"] ,pr_ad.msg_ad["fail_exception"])
                
            else:
                cm.xlsx(pr_ad.ad_basic["fail"] ,pr_ad.msg_ad["fail_exception_add"])

        if  check_apex == True :
            driver.refresh()
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,data["iframe_vc"])))
            driver.switch_to.frame(driver.find_element_by_xpath(data["iframe_vc"]))
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,data["menu_vc"]))).click()
            time.sleep(3)

            driver.find_element_by_link_text("Request Vacation").click()
            button_select_approval = cm.is_Displayed(data["rq_vc"]["bt_select_approver"])
            if button_select_approval == True :
                cm.xlsx(pr_ad.ad_basic["fail"] ,pr_ad.msg_ad["fail_exception_suc"])
            else:
                cm.xlsx(pr_ad.ad_basic["pass"] ,pr_ad.msg_ad["pass_exception_suc"])

            ip_use_ex_approval = driver.find_element_by_xpath(pr_ad.dt_ad["ip_use_ap_exception"])
            driver.find_element_by_xpath(pr_ad.dt_ad["use_ap_exception"]).click()

            time.sleep(3)
            use_ex_approval_after = ip_use_ex_approval.is_selected()
            if use_ex_approval_before != use_ex_approval_after :
                WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,data["iframe_vc"])))
                driver.switch_to.frame(driver.find_element_by_xpath(data["iframe_vc"]))
                WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,data["menu_vc"]))).click()

                time.sleep(3)
                driver.find_element_by_link_text("Request Vacation").click()
                button_select_approval = cm.is_Displayed(data["rq_vc"]["bt_select_approver"])
                if button_select_approval == True :
                    cm.xlsx(pr_ad.ad_basic["fail"] ,pr_ad.msg_ad["fail_exception_suc"])
                else:
                    cm.xlsx(pr_ad.ad_basic["pass"] ,pr_ad.msg_ad["pass_exception_suc"])
    except:
        driver.find_element_by_xpath(data["sm_basic_settings"]).click()

def sm_bs_delete_approval_exception():
    i = j = 1
    result_delete = False
    removed = True


    time.sleep(2)
    list_apex=driver.find_elements_by_xpath(pr_ad.dt_ad["list_apex"])
    total_apex_before = cm.total_data(list_apex)
    
    if total_apex_before == 0 :
        cm.xlsx(pr_ad.ad_delete_ax["pass"] ,pr_ad.msg_ad["pass_delete_no_ex"])
    else:
        ax_to_delete = driver.find_element_by_xpath(pr_ad.dt_ad["apex3"]).text 

        driver.find_element_by_css_selector(pr_ad.dt_ad["ic_delete_ax"]).click()
        if cm.is_Displayed(pr_ad.dt_ad["bt_close_ar"])  == True : 

            cm.msg("p",pr_ad.msg_ad["pass_delete_ex_icon"])
            driver.find_element_by_xpath(pr_ad.dt_ad["bt_dele_ar"]).click()

            if cm.is_Displayed(pr_ad.dt_ad["bt_close_ar"]) == False :
                cm.xlsx(pr_ad.ad_delete_ax["pass"] ,pr_ad.msg_ad["pass_delete_ex"])
                result_delete = True
                
            else:
                cm.xlsx(pr_ad.ad_delete_ax["fail"] ,pr_ad.msg_ad["fail_delete_ex"])
                
        else:
            cm.msg("f",pr_ad.msg_ad["fail_delete_ex_icon"])


        if result_delete == True :
            time.sleep(2)
            list_apex = driver.find_elements_by_xpath(pr_ad.dt_ad["list_apex"])
            total_apex_after = cm.total_data(list_apex)

            if total_apex_after == 0 :
                if total_apex_before == 1 :
                    cm.xlsx(pr_ad.ad_delete_ax["pass"] ,pr_ad.msg_ad["pass_delete_ex_removed"])
                else:
                    cm.xlsx(pr_ad.ad_delete_ax["fail"] ,pr_ad.msg_ad["fail_delete_ex_removed"])

            else :
                while j <= total_apex_after:
                    name_ax = driver.find_element_by_xpath(xpath.ax_name(j)).text
                    if ax_to_delete == name_ax :
                        cm.xlsx(pr_ad.ad_delete_ax["fail"] ,pr_ad.msg_ad["fail_delete_ex_removed"])
                        removed = False
                        break
                    j = j+1
                if removed == True :
                    cm.xlsx(pr_ad.ad_delete_ax["pass"] ,pr_ad.msg_ad["pass_delete_ex_removed"])
         
def submenu_create_vacation():
    
    cm.msg("n", "SUB MENU : ADMIN SETTINGS ")
    cm.msg("n", "I.Create Vacation")
    
    
    
    driver.find_element_by_link_text("Create Vacation").click()
    if cm.is_Displayed(pr_ad.dt_ad["tab_list_vc"]) == True :
        cm.xlsx(pr_ad.ad_sub_menu["pass"] ,pr_ad.msg_ad["pass_access_sb_create_vc"])
        
        cm.msg("n", "Create Vacation")
        sm_cre_vc_create_vacation()
        
        
        cm.msg("n", "Tab Vacation List")
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,pr_ad.dt_ad["tab_list_vc"]))).click()
        if cm.is_Displayed(pr_ad.dt_ad["text_list_vc"]) == True :

            cm.xlsx(pr_ad.ad_tab_create["pass"] ,pr_ad.msg_ad["pass_access_tb_vc_li"])

            cm.msg("n", "Detail Vacation")
            sm_cre_vc_view_detail_vacation()

            cm.msg("n", "Search Vacation")
            sm_cre_vc_search_vacation()
            
            cm.msg("n", "Delete vacation")
            sm_cre_vc_delete_vacation()
            

            cm.msg("n", "Filter Vacation")
            sm_cre_vc_filter_vacation()
            
        else:
            cm.xlsx(pr_ad.ad_tab_create["fail"] ,pr_ad.msg_ad["fail_access_tb_vc_li"])
        
        
        cm.msg("n", "Tab Create Vacation")
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,pr_ad.dt_ad["tab_create_vc"]))).click()
        if cm.is_Displayed(pr_ad.dt_ad["text_create_vc"]) == True :
            cm.xlsx(pr_ad.ad_tab_create["pass"],pr_ad.msg_ad["pass_access_tb_create_vc"])
        else:
            cm.xlsx(pr_ad.ad_tab_create["fail"] ,pr_ad.msg_ad["fail_access_tb_create_vc"])
        
    else:
        cm.xlsx(pr_ad.ad_sub_menu["fail"] ,pr_ad.msg_ad["fail_access_sb_create_vc"])

def submenu_manager_settings():
   
    cm.msg("n", "II.Manager Settings")
    driver.find_element_by_xpath(data["sm_manager_settings"]).click()
    if cm.is_Displayed(pr_ad.dt_ad["bt_add_manager"]) ==True :
        cm.xlsx(pr_ad.ad_sub_mn["pass"] ,pr_ad.msg_ad["pass_access_sb_manager_st"])
        
        
        cm.msg("n", "Add Manager")
        sm_mn_se_add_manager()
        
        cm.msg("n", "Search Manager")
        sm_mn_se_search_manager()
        
        
        cm.msg("n","Modify Manager'permission")
        sm_mn_se_modify_permission()
        
        
        cm.msg("n","Delete manager")
        sm_mn_se_delete_manager()
        
        
        cm.msg("n", "Tab Approval Setting")
        driver.find_element_by_xpath(pr_ad.dt_ad["mn_approval_settings"]).click()
        if cm.is_Displayed(pr_ad.dt_ad["bt_select_approver"]) == True :
            cm.xlsx(pr_ad.ad_tab_ap["pass"] ,pr_ad.msg_ad["pass_access_tb_approval_st"])
            
            cm.msg("n", "Add user is Arbitrary Decision ")
            sm_mn_se_add_arbitrary_decision()
            
            cm.msg("n", "Delete user is Arbitrary Decision ")
            sm_mn_se_delete_arbitrary_decision()
            
        else :
            cm.xlsx(pr_ad.ad_tab_ap["fail"] ,pr_ad.msg_ad["fail_access_tb_approval_st"])
       
        
        cm.msg("n", "Tab Manager Settings")
        driver.find_element_by_xpath(pr_ad.dt_ad["tab_mn_settings"]).click()
        if cm.is_Displayed(pr_ad.dt_ad["text_mn_settings"]) == True :
            cm.xlsx(pr_ad.ad_tab_ap["pass"] ,pr_ad.msg_ad["pass_access_tb_manager_st"])
        else:
            cm.xlsx(pr_ad.ad_tab_ap["fail"] ,pr_ad.msg_ad["fail_access_tb_manager_st"])
        
    else:
        cm.xlsx(pr_ad.ad_sub_mn["fail"] ,pr_ad.msg_ad["fail_access_sb_manager_st"])

def submenu_basic_settings():
    
    cm.msg("n", "III.Basic Settings")
    

    driver.find_element_by_xpath(data["sm_basic_settings"]).click()
    if cm.is_Displayed(pr_ad.dt_ad["bt_add_approval_exception"]) == True :
        cm.xlsx(pr_ad.ad_sub_bs["pass"] ,pr_ad.msg_ad["pass_access_sb_basic"])
        

        cm.msg("n", "Use Settlement")
        sm_bs_use_settlement()
        
        cm.msg("n", "Use Hour unit")
        sm_bs_use_hour_unit()

        cm.msg("n", "Use vacation schedule")
        sm_bs_use_vacation_schedule()
        
        cm.msg("n", "Use hour-unit limit")
        sm_bs_use_hour_unit_limit()
        
       
        cm.msg("n", "Add Approval Exception")
        sm_bs_use_approval_exception()

        
        cm.msg("n", "Delete Approval Exception")
        sm_bs_delete_approval_exception()  

    else:
        cm.xlsx(pr_ad.ad_sub_bs["fail"] ,pr_ad.msg_ad["fail_access_sb_basic"])

def submenu_manage_history():
    driver.find_element_by_link_text("Manager History").click()
    if cm.is_Displayed(pr_ad.dt_ad["text_sub_mn_history"]) == True :
        cm.xlsx(pr_ad.ad_sub_his["pass"] ,pr_ad.msg_ad["pass_access_sb_manager_his"])
    else:
        cm.xlsx(pr_ad.ad_sub_his["fail"] ,pr_ad.msg_ad["fail_access_sb_manager_his"])

def admin_setting():
    if cm.is_Displayed1("textlink","Create Vacation") == True:
        submenu_create_vacation()
        submenu_manager_settings()
        submenu_basic_settings()
        submenu_manage_history()
    else:
        cm.xlsx(pr_ad.ad_permission["pass"] ,pr_ad.msg_ad["pass_not_admin"])

   
    
   
   
    

 
    


