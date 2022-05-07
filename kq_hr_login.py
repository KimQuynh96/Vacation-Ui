
from kq_hr_lib_common import driver,data,cm,Keys,EC,By,WebDriverWait,json
from kq_hr_param import pr_log 

def login(domain):
    cm.msg("n","LOGIN")
    url = json.loads(cm.param_url(domain))
    driver.get("http://"+domain+"/ngw/app/#/sign")
    
     
    driver.implicitly_wait(10)
    driver.find_element_by_id("log-userid").send_keys(data["user"])
    cm.msg("p" ,pr_log.msg_log["input_id"])
   
    driver.switch_to.frame(driver.find_element_by_id("iframeLoginPassword"))
    driver.find_element_by_id("p").send_keys(data["pass"])
    driver.switch_to.default_content()
    cm.msg("p" ,pr_log.msg_log["input_pass"])

    driver.find_element_by_id("btn-log").send_keys(Keys.RETURN)
    cm.msg("p" ,pr_log.msg_log["login"])
    
    if cm.is_Displayed1("id" ,"menubar") == True:
        cm.xlsx(pr_log.log["pass"] ,pr_log.msg_log["pass_login"])
        pr_log.result_access["login"] = True
        
        
    else:
        cm.xlsx(pr_log.log["fail"] ,pr_log.msg_log["fail_login"])
        pr_log.result_access["login"] = False

    if  pr_log.result_access["login"] == True:
        driver.get(url["menu_vacation"])
        
        #driver.get("http://qavn1.hanbiro.net/nhr/hr/timecard/dashboard?mode=user")
        #time.sleep(20)
        WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH,data["iframe_vc"])))
        driver.switch_to.frame(driver.find_element_by_xpath(data["iframe_vc"]))
        
        if cm.is_Displayed(data["menu_vc"]) == True :
            driver.find_element_by_xpath(data["menu_vc"]).click()

            if cm.is_Displayed(data["sm_my_vc_sta"]) == True :
                cm.msg("n" ,"ACCESS VACATION")
                pr_log.result_access["access_menu"] = True
                cm.xlsx(pr_log.menu["pass"] ,pr_log.msg_log["pass_access"])
                
                cm.popup_time_card()

            else:
                pr_log.result_access["access_menu"] = False
                cm.xlsx(pr_log.menu["fail"] ,pr_log.msg_log["fail_access"])

        else:
            pr_log.result_access["access_menu"] = False
            cm.xlsx(pr_log.menu["fail"] ,pr_log.msg_log["fail_access_hr"])
            
    return pr_log.result_access
   
