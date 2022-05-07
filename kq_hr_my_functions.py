from kq_hr_lib_common import datetime,json,re,time,pipe,EC,By,WebDriverWait,Keys,relativedelta
from kq_hr_param import pr_rq ,pr_mn,submenu_my_vacation
from kq_hr_lib_my_vacation import cm_rq,cm_fu,xpath,all_day,consecutive,half_day,hour_unit,cancel,used,filter
from kq_hr_lib_common import cm , data , driver 
param = json.loads(submenu_my_vacation())



def request_vacation_all_day(request_date,result_select_approver,approver):
    
    full_day = all_day.request_all_day(request_date,approver)
    all      = all_day(full_day)
    xpath.send_request()
    if  bool(all.result_select_vc)       == True and \
        bool(result_select_approver)     == True and \
        bool(all.result_date)            == True     :

        cm_fu.click_on_button_to_request()
        result_request = cm_rq.check_result_request()
        if result_request == "pass":
            
            cm_rq.information_vacation(pr_rq.msg_re['info_vc'],all.vacation_request)
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_ad_request"])
            cm_rq.created_request_and_view_detail(all.vacation_request,"all",all.use_hour_unit,approver)
            number_after_request  = cm_rq.available_vacation()
            number_before_request = xpath.par_number_of_days(all,number_after_request,"all")
            cm_rq.check_number_of_days_off(**number_before_request)
            cm_rq.vacation_displayed_in_time_card(all.vc_date)
            cancel.cancel_request(all.oneday,all.use_hour_unit)
        
        elif result_request == "fail":
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re['pass_ad_cannot_request'])

        else:
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_ad_not_conditions"])
            cm.msg("p" , "-Msg :",result_request)
    
def request_vacation_consecutive(request_date,result_select_approver,approver):

    conse = consecutive.vacation_consecutive(request_date,approver)
    con   = consecutive(conse)
    xpath.send_request()
    if  bool(con.result_select_vc)   == True  and \
        bool(result_select_approver) == True  and \
        bool(con.result_date)        == True      :

        cm_fu.click_on_button_to_request()
        result_request = cm_rq.check_result_request()
        if result_request == "pass":
            cm_rq.information_vacation(pr_rq.msg_re['info_vc'],con.vacation_request)
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_cs_request"])
            cm_rq.created_request_and_view_detail(con.vacation_request,"vc_con",con.use_hour_unit,approver)
            number_after_request = cm_rq.available_vacation()
            number_before_request = xpath.par_number_of_days(con,number_after_request,"vc_con")
            cm_rq.check_number_of_days_off(**number_before_request)
            cancel.cancel_request(con.oneday,con.use_hour_unit)
        
        elif result_request == "fail":
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_cs_cannot_request"])

        else:
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_cs_not_conditions"])
            cm.msg("p" ,"-Msg :",result_request)
    
def request_vacation_half_am(request_date,result_select_approver,approver):
   
    half = half_day.vacation_half(request_date,approver,"am")
    ha   = half_day(half)
    xpath.send_request()
    if  bool(ha.result_select_vc)     != False and \
        bool(result_select_approver)  == True  and \
        bool(ha.result_date)          == True      :
        cm_fu.click_on_button_to_request()
        result_request = cm_rq.check_result_request()
        if result_request == "pass":
            cm_rq.information_vacation(pr_rq.msg_re['info_vc'],ha.vacation_request)
            cm.xlsx( pr_rq.re_vc["pass"] ,pr_rq.msg_re["pass_am_request"])

            '''
            msg("t" ,"Time from time card :" + time_time_card)
            msg("t" ,"Time from request vacation :" + time_request)
            '''
            cm_rq.created_request_and_view_detail(ha.vacation_request,"half_day",ha.use_hour_unit,approver)
            number_after_request = cm_rq.available_vacation()
            number_before_request = xpath.par_number_of_days(ha,number_after_request,"half_day")
            cm_rq.check_number_of_days_off(**number_before_request)
            cancel.cancel_request(ha.oneday,ha.use_hour_unit)
            
        elif result_request == "fail":
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_am_cannot_request"])
        else:
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_am_not_conditions"])
            cm.msg("p" ,"-Msg :",result_request)
    
def request_vacation_half_pm(request_date,result_select_approver,approver):

    half = half_day.vacation_half(request_date,approver,"pm")
    ha   = half_day(half)
    xpath.send_request()
    if  bool(ha.result_select_vc)    != False and \
        bool(result_select_approver) == True  and \
        bool(ha.result_date)         == True      :
        cm_fu.click_on_request_button()
        result_request =  cm_rq.check_result_request()
        if result_request == "pass":
            cm_rq.information_vacation(pr_rq.msg_re['info_vc'],ha.vacation_request)
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_pm_request"])
            cm_rq.created_request_and_view_detail(ha.vacation_request,"half_day",ha.use_hour_unit,approver)
            number_after_request =  cm_rq.available_vacation()
            number_before_request = xpath.par_number_of_days(ha,number_after_request,"half_day")
            cm_rq.check_number_of_days_off(**number_before_request)
            cancel.cancel_request(ha.oneday,ha.use_hour_unit)
           
        elif result_request == "fail":
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_pm_cannot_request"])

        else:
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_pm_not_conditions"])
            cm.msg("p" ,"-Msg :",result_request)
    
def request_vacation_hour_unit(request_date,result_select_approver,approver):
    
    hour = hour_unit.vacation_hour_unit(request_date,approver,"hour")
    hu   = hour_unit(hour)
    xpath.send_request()
    if  bool(result_select_approver) == True and \
        bool(hu.result_date)         == True and \
        bool(hu.result_select_vc)    != False    :
        
        cm_fu.click_on_request_button()
        result_request =  cm_rq.check_result_request()
        if result_request == "pass":
            cm_rq.information_vacation(pr_rq.msg_re['info_vc'],hu.vacation_request)
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_hu_request"])
            cm_rq.created_request_and_view_detail(hu.vacation_request,"hour",hu.use_hour_unit,approver)
            number_after_request =  cm_rq.available_vacation()
            number_before_request = xpath.par_number_of_days(hu,number_after_request,"hour")
            cm_rq.check_number_of_days_off(**number_before_request)
            cancel.cancel_request(hu.oneday,hu.use_hour_unit)

        elif result_request == "fail":
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_hu_cannot_request"])
            
        else:
            cm.xlsx( pr_rq.re_vc["pass"] , pr_rq.msg_re["pass_hu_not_conditions"])
            cm.msg("p" ,"-Msg :",result_request)

def request_vacation_on_used_date(request_date,result_select_approver,approver):
    used.list_vc()

def delete_request_my_vc():
    try :
        i = 1
        result_delete = result_find =False
        if cm.is_Displayed(pr_rq.rq_vc["check_list_re"]) == bool(True) :
            cm.xlsx(pr_rq.re_delete["pass"] ,pr_rq.msg_re["pass_no_delete"])
        else:
            # Find request to delete #
            total_request_before_delete = cm_rq.count_all_vacation_request()
            time.sleep(3)
            list_request = driver.find_elements_by_xpath(pr_rq.rq_vc["list_request"])
            total_request_before = cm.total_data(list_request)
            while i<= total_request_before:
                if cm.is_Displayed(cm.xpath2(pr_rq.rq_vc["de_ic"],str(i-1),"')]")) == bool(True) : 
                        result_find= True
                        driver.find_element_by_xpath(cm.xpath2(pr_rq.rq_vc["de_ic"],str(i-1),"')]")).click()

                        if cm.is_Displayed(pr_rq.rq_vc["bt_cl_re"]) == bool(True) :  

                            cm.msg("p",pr_rq.msg_re["pass_delete_click"])
                            driver.find_element_by_xpath(pr_rq.rq_vc["bt_de_re"]).click()
                            time.sleep(1)
                            if cm.is_Displayed(pr_rq.rq_vc["bt_cl_re"]) == bool(False) : 
                                cm.xlsx(pr_rq.re_delete["pass"] ,pr_rq.msg_re["pass_delete"])
                                result_delete=True
                            else:
                                cm.xlsx(pr_rq.re_delete["fail"] ,pr_rq.msg_re["fail_delete"])
                        
                        else:
                            cm.xlsx(pr_rq.re_delete["fail"] ,pr_rq.msg_re["fail_delete_click"])
                        break 
                i=i+1 
            
            # If the request is deleted, check the request removed from request list #
            if result_delete == True:
                i = 1
                total_request_after_delete = cm_rq.count_all_vacation_request()
                if total_request_before_delete - 1 == total_request_after_delete:
                    cm.xlsx(pr_rq.re_delete["pass"] ,pr_rq.msg_re["pass_request_removed"])
                else:
                    cm.xlsx(pr_rq.re_delete["fail"] ,pr_rq.msg_re["fail_request_removed"])

            if result_find== False:
                cm.xlsx(pr_rq.re_delete["pass"] ,pr_rq.msg_re["pass_no_delete"])
    except:
        driver.find_element_by_link_text("My Vacation Status").click()


def filter_status(type):
    try:
        if cm.is_Displayed(pr_rq.rq_vc["check_list_re"]) == True :
            cm.xlsx(pr_rq.re_filter["pass"] ,pr_rq.msg_re["pass_no_filter"])
        else:
            list_status_before = filter.list_status_before(type)

            # < Filter by status is Request #>
            result_request = filter.filter_status(**xpath.par_request(list_status_before))

            # < Filter by status is Progressing #>
            result_progressing = filter.filter_status(**xpath.par_progress(list_status_before))

            # < Filter by status is Completed #>
            result_completed = filter.filter_status(**xpath.par_completed(list_status_before))
            
            if result_request == True and result_progressing == True and result_completed == True:
                cm.xlsx(pr_rq.re_filter["pass"] ,pr_rq.msg_re["pass_filter"])
            else:
                cm.xlsx(pr_rq.re_filter["fail"] ,pr_rq.msg_re["fail_filter"])
    except:

        if type != "cc" :
            driver.find_element_by_link_text("My Vacation Status").click()
        else :
            driver.find_element_by_link_text("View CC").click()

def view_detail_request_cc():
    
    try :
        result_vc_date = result_number_use = result_date = False
        if cm.is_Displayed(pr_rq.rq_vc["check_list_re"]) == bool(True) :
            cm.xlsx(pr_rq.re_view_cc["pass"] ,pr_rq.msg_re["pass_view_detail"])
        else:
            driver.find_element_by_xpath(pr_rq.rq_vc["detail_cc"]).click()
            if cm.is_Displayed(pr_rq.rq_vc["check_cc"]) == bool(True):
                cm.msg("p",pr_rq.msg_re["pass_view_click"])

                vc_date_title = driver.find_element_by_xpath(pr_mn.mn_pro["content_request"]).text
                vc_date_content = driver.find_element_by_xpath(pr_mn.mn_pro["content_request1"]).text 
                if vc_date_title.strip() == "Vacation Date" and len(vc_date_content.strip()) != 0 :
                    cm.msg("p",pr_rq.msg_re["pass_view_date"] )
                    result_vc_date=True
                else:
                    cm.xlsx(pr_rq.re_view_cc["fail"] ,pr_rq.msg_re["fail_view_date"])

                use_title = driver.find_element_by_xpath(pr_mn.mn_pro["content_use"]).text
                use_content = driver.find_element_by_xpath(pr_mn.mn_pro["content_use1"]).text 
                if use_title.strip() == "Use" and len(use_content.strip()) != 0 :
                    cm.msg("p",pr_rq.msg_re["pass_number_user"])
                    result_number_use=True
                else:
                    cm.xlsx(pr_rq.re_view_cc["fail"] ,pr_rq.msg_re["fail_number_user"])
                
            
                request_date_title = driver.find_element_by_xpath(pr_mn.mn_pro["content_re_date"]).text
                request_date_content = driver.find_element_by_xpath(pr_mn.mn_pro["content_re_date1"]).text
                if request_date_title.strip() =="Request date" and len(request_date_content.strip()) !=0:
                    cm.msg("p",pr_rq.msg_re["pass_view_request_date"])
                    result_date=True
                else:
                    cm.xlsx(pr_rq.re_view_cc["fail"] ,pr_rq.msg_re["fail_view_request_date"])
                

                if result_date == True and result_number_use ==True and result_vc_date == True:
                    cm.xlsx(pr_rq.re_view_cc["pass"] ,pr_rq.msg_re["pass_view"])
                else:
                    cm.xlsx(pr_rq.re_view_cc["fail"] ,pr_rq.msg_re["fail_view"])

                driver.find_element_by_link_text("View CC").click()

            else:
                cm.xlsx(pr_rq.re_view_cc["fail"] ,pr_rq.msg_re["fail_view_click"])
    except:
        driver.find_element_by_link_text("View CC").click()


def request_and_cancel_vacation():

    cm.msg("n" , "I.REQUEST VACATION")
    request_date = str(datetime.date.today())
   
        
    cm.msg("n" ,"Select approver")
    result_select_approver = False
    approver = cm_rq.select_approver()
    if approver["result_approver"] == True:
        result_select_approver = True
    else:
        if approver["approval_line"] == True:
            result_select_approver = True
        if approver["approval_exception"] == True:
            result_select_approver = True

    driver.find_element_by_link_text("Request Vacation").click()
    total_vc = cm_rq.total_vacation()
    
    if total_vc == 0 :
        cm.xlsx( pr_rq.re_request["pass"] , pr_rq.msg_re["pass_no_vacation_request"])

    else:
        
        cm.msg("n", "Request Vacation : All Day ")
        request_vacation_all_day(request_date,result_select_approver,approver)
        
        cm.msg("n", "Request Vacation : Consecutive Vacation")
        request_vacation_consecutive(request_date,result_select_approver,approver)
        
        
        cm.msg("n", "Request Vacation : Half day(Am) ")
        request_vacation_half_am(request_date,result_select_approver,approver)
        
        
        cm.msg("n", "Request Vacation : Half day(Pm)")
        request_vacation_half_pm(request_date,result_select_approver,approver)
        

        cm.msg("n",  "Request Vacation : Hour Unit")
        request_vacation_hour_unit(request_date,result_select_approver,approver)
        
        #request_vacation_on_used_date(request_date,result_select_approver,approver)
    

'''
# linux #
def request():   
    result_access_menu_vacation = access_menu_vacation()
    if result_access_menu_vacation ==  True :
        submenu_request_vacation()
'''

# Window #
def request(domain):   
    cm_rq.login(domain)
    result_access_menu_vacation = cm_rq.access_menu_vacation(domain)
    if result_access_menu_vacation ==  True :
        request_and_cancel_vacation()
       

    

    
    