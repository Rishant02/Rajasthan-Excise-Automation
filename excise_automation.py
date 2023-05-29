import os
import pymongo
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from send_mail import send_mail
import win32com.client as win32
from generate_pivot import run_all_excel,run_rkl_excel
from upload import upload_root_file
import pandas as pd
from datetime import datetime,timedelta
from dotenv import load_dotenv
load_dotenv()
win32c = win32.constants

class ExciseAutomation:
    def __init__(self) -> None:
        self.conn_string = os.getenv('MONGO_URI')
        self.client = pymongo.MongoClient(self.conn_string)
        self.db = self.client['rajasthan']
        self.download_path = os.path.join(os.getcwd(),'static')
        self.yesterday = datetime.now() + timedelta(days=-1)
        self.sheet_name = '01-{} (RSBCL)'.format(self.yesterday.strftime('%d %b'))
        self.sheet_name_email = '01-{}'.format(self.yesterday.strftime('%d %b-%Y'))
        self.from_email = os.getenv('SMTP_EMAIL_ADDRESS')
        # self.to = ['yadavrd@radico.co.in']
        # self.cc = ['mastwalrk@radico.co.in','yogeshk@radico.co.in','kapoorm@radico.co.in','barupalmk@radico.co.in']
        self.to = ['mastwalrk@radico.co.in']
        self.cc = ['rishantmastwal@gmail.com']
        self.password = os.getenv('SMTP_PASSWORD')
        self.subject = 'AUTOMATED RSBCL REPORT FROM {} (January Master)'.format(self.sheet_name_email)
        self.file_path = f'rsbcl_report {self.sheet_name_email}.xlsx'
        self.ONE_DRIVE_FOLDER = 'https://radicokhaitan-my.sharepoint.com/:f:/g/personal/mastwalrk_radico_co_in/EuO1SPZSe3lJpf2Jq7sWgKUBAqztT6xIszYUndfQM2_AaQ?e=xZzriX'
        # Getting master data stored in mongodb
        self.packings_df = self.retrieve_mongodb_data_as_df(self.db['packings'],{'_id':0,'__v':0})
        self.deo_df=self.retrieve_mongodb_data_as_df(self.db['deo_offices'],{'_id':0,'__v':0})
        self.brand_pack_df=self.retrieve_mongodb_data_as_df(self.db['brand_packs'],{'_id':0,'__v':0})
        self.brand_pack_df.rename(columns={"BRAND_NAME_WITH_PACKING":"BRAND WITH PACKING","BRAND_NAME":"BRAND NAME"},inplace=True)
        self.rkl_brands_df=self.retrieve_mongodb_data_as_df(self.db['rkl_brands'],{'_id':0,'__v':0})
        self.rkl_brands_df.rename(columns={"RKL_Brand_Full_Name":"BRAND WITH PACKING","RKL_Brand_Short_Name":"RKL Brand Name"},inplace=True)
        self.rkl_brands=sorted(set(self.rkl_brands_df['RKL Brand Name']))
        self.companies_df=self.retrieve_mongodb_data_as_df(self.db['brand_companies'],{'_id':0,'__v':0,'SEGMENT':0,'Category':0})
        self.companies_df.rename(columns={"Brands":"BRAND NAME","Company_Name":"Company Name"},inplace=True)
        keys=list(self.db['licensee_group_january_2023'].find_one().keys())+['Secondary_Owner','Secondary_Contact','K1','K2','K3']
        to_remove=['Licensee_Name','Vends','Location','Group_Name','Category','Owner_Name']
        self.licensee_df=self.retrieve_mongodb_data_as_df(self.db['licensee_group_january_2023'],{k:0 for k in keys if k not in to_remove})
        self.licensee_df.rename(columns={'Licensee_Name':'LICENSEE_NAME','Vends':'Vend','Group_Name':'Group Name','Category':'Category Group/Ind.','Owner_Name':'Name Of Owner/Retailer'},inplace=True)
        self.groupwise_license=self.retrieve_mongodb_data_as_df(self.db['licensee_group_january_2023'],{'_id':0,'__v':0})
        self.groupwise_license['RSBCL_CODE']=self.groupwise_license['RSBCL_CODE'].astype(str)
        self.groupwise_license['RSBCL_CODE']=self.groupwise_license['Licensee_Name'].str.split('-').str[-1]
        self.groupwise_license.rename(columns={'Licensee_Name':'LICENSEE_NAME'},inplace=True)

    def retrieve_mongodb_data_as_df(self, collection, filter_col:dict[str,int])->pd.DataFrame:
        data = list(collection.find({},filter_col))
        return pd.DataFrame(data).drop_duplicates()
    
    def login_user(self,driver,username,password):
        driver.get(os.getenv('WEBSITE_URL'))
        close_modal = driver.find_element(By.XPATH,'//*[@id="ModalHome"]/div/div/div[3]/button')
        close_modal.click()
        login_anchor=driver.find_element(By.XPATH,'//*[@id="div_menu"]/ul/li[5]/a')
        login_anchor.click()
        rsbcl_btn=driver.find_element(By.XPATH,'//*[@id="div_menu"]/ul/li[5]/ul/li[2]/a')
        rsbcl_btn.click()
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="txt_userid"]')))
        username_input=driver.find_element(By.XPATH,'//*[@id="txt_userid"]')
        password_input=driver.find_element(By.XPATH,'//*[@id="txt_pass"]')
        captcha_el=driver.find_element(By.ID,'mainCaptcha')
        captcha_input=driver.find_element(By.ID,'txtInput')
        captcha_text=captcha_el.text
        username_input.send_keys(username)
        password_input.send_keys(password)
        captcha_input.send_keys(captcha_text)
        login_btn=driver.find_element(By.XPATH,'//a[@id="btnlsubmit"]')
        login_btn.click()
        # To avoid change password screen if appears
        try:
            cancel_btn=driver.find_element(By.ID,'btnCancel')
            cancel_btn.click()
        except Exception as e:
            print('Cancel button is not found.')
        driver.implicitly_wait(2)

    def download_data_to_dataframe(self,driver)->pd.DataFrame:
        WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.ID,'form1')))
        new_reports=driver.find_element(By.XPATH,'//input[@id="ModulesList_ctl02_iBtnModules"]')
        new_reports.click()
        redirect_btn=driver.find_element(By.ID,'btnRedirect')
        redirect_btn.click()

        reports_anchor=driver.find_element(By.XPATH,'//a[text()="Reports"]')
        reports_anchor.click()
        sales_reports_anchor=WebDriverWait(driver,20).\
            until(EC.element_to_be_clickable((driver.find_element(By.XPATH,'//a[text()="sales reports"]'))))
        sales_reports_anchor.click()
        licensee_export_form=driver.find_element(By.XPATH,'//a[text()="Licensee Wise Sales(Export To Excel)"]')
        licensee_export_form.click()

        from_date=WebDriverWait(driver,20).until(EC.presence_of_element_located((By.ID,'ctl00_cph1_txtfromdate')))
        from_date.send_keys(self.yesterday.strftime(r'01/%m/%Y'))
        # from_date.send_keys('01/03/2023')

        to_date=driver.find_element(By.ID,'ctl00_cph1_txttodate')
        to_date.send_keys(self.yesterday.strftime(r'%d/%m/%Y'))
        # to_date.send_keys('31/03/2023')  

        group_code=Select(driver.find_element(By.ID,'ctl00_cph1_ddlgroupcode'))
        group_code.select_by_visible_text('IMFL')

        export_btn=driver.find_element(By.ID,'ctl00_cph1_btnshow')
        export_btn.click()

        licensee_file_path = os.path.join(self.download_path,'Licensee Wise Sales Report For Supplier .xls')
        try:
            WebDriverWait(driver,600).until(
                lambda x:os.path.exists(licensee_file_path)
            )
        except TimeoutException as e:
            raise TimeoutException('Downloading raw data has taken more than 10 minutes. This may be due to a problem with the website or a high volume of traffic.\nThe bot has attempted to download the data three times without success.')
        driver.quit()
        print('Doing computations on downloaded file')
        html_file=pd.read_html(licensee_file_path)
        os.remove(licensee_file_path)
        df=html_file[0]
        df.columns=df.iloc[0]
        df=df[1:]
        # Some cleaning
        type_dict={'TOTAL_CASE':int,'TOTAL_BTL':int,'TOTAL_BULK_LITER':float}
        df=df.astype(type_dict)
        df=df.rename(columns={'BRAND':'BRAND WITH PACKING'})
        df['BRAND WITH PACKING']=df['BRAND WITH PACKING'].str.replace(r'\x80\x99','™')
        df['BRAND WITH PACKING']=df['BRAND WITH PACKING'].str.replace(r'\x80\x93','€')
        df['LICENSEE_NAME']=df['LICENSEE_NAME'].str.upper()

        return df
    def transform_df_by_master_and_export(self,df:pd.DataFrame)->str:
        merged_df=pd.merge(df,self.packings_df,on="PACKING_IN_ML",how='left')
        # merged_df['Cases']=round(merged_df['TOTAL_CASE']+merged_df['TOTAL_BTL']/merged_df['PACKING'],3)
        merged_df['Cases']=merged_df['TOTAL_CASE']+merged_df['TOTAL_BTL']/merged_df['PACKING']
        merged_df=merged_df.drop('PACKING',axis=1)

        final_df=pd.merge(pd.merge(pd.merge(pd.merge(pd.merge(merged_df,self.rkl_brands_df,on='BRAND WITH PACKING',how='left').fillna('Other'),
                                            self.brand_pack_df,on='BRAND WITH PACKING',how='left'),
                                            self.companies_df,on='BRAND NAME',how='left').fillna('Other'),
                                            self.deo_df,on='DEO_OFFICE_NAME',how='left'),
                                            self.licensee_df,on='LICENSEE_NAME',how='left')
        rkl_df=final_df[final_df['RKL Brand Name']!='Other']
        print('Starting rkl groupwise')
        merged_group_df=pd.merge(self.groupwise_license,rkl_df,on='LICENSEE_NAME',how='left')
        pivoted_df = merged_group_df.pivot_table(index='LICENSEE_NAME', columns='RKL Brand Name', values='Cases', aggfunc='sum', fill_value=0)
        pivoted_cols=pivoted_df.columns
        for brand in self.rkl_brands:
            if brand not in pivoted_cols:
                pivoted_df[brand] = 0
        pivoted_df=pivoted_df.reset_index()
        groupwise_rkl=pd.merge(self.groupwise_license,pivoted_df,on='LICENSEE_NAME',how='left')
        groupwise_rkl.iloc[:,18:]=groupwise_rkl.iloc[:,18:].fillna(0)
        writer=pd.ExcelWriter(self.file_path)
        print('Exporting converted file...')
        final_df.to_excel(writer,index=False,sheet_name=self.sheet_name)
        groupwise_rkl.to_excel(writer,index=False,sheet_name='groupwise_rkl')
        writer.close()
        excel=win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible=False

        pt_fields=[[col,f'Sum of {col}',win32c.xlSum,'0.000'] for col in pivoted_cols]
        wb = excel.Workbooks.Open(os.path.abspath(os.getcwd()+f'\\{self.file_path}'))
        run_all_excel(excel,wb)
        print('rkl pivot table starting...')
        run_rkl_excel(excel,wb,pt_fields)
        wb.Close(SaveChanges=1)
        excel.Application.Quit()
        return self.file_path
    
    def run_driver(self):
        prefs = {
        "download.default_directory": self.download_path,
        "download.prompt_for_download": False,
        # "download.directory_upgrade": True,
        # "safebrowsing.enabled": False
        }
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        chrome_options.add_experimental_option('prefs',prefs)
        chrome_options.add_argument('--incognito')
        driver= webdriver.Chrome(options=chrome_options)
        driver.implicitly_wait(2)
        driver.maximize_window()
        try:
            self.login_user(driver,os.getenv('WEBSITE_USERNAME'),os.getenv('WEBSITE_PASSWORD'))
            df = self.download_data_to_dataframe(driver)
            generated_file_path = self.transform_df_by_master_and_export(df)
            res = upload_root_file(generated_file_path)
            body = ''
            if res:
                body='''
                    <html>
                        <head>
                            <style>
                                body{{
                                    'font-family: Arial, sans-serif;'
                                    'background-color: #f2f2f2;'
                                }}
                                .container{{
                                    'width: 80%;'
                                    'margin: 0 auto;'
                                    'padding: 20px;'
                                    'background-color: #fff;'
                                    'border-radius: 10px;'
                                    'box-shadow: 0px 2px 5px rgba(0, 0, 0, 0.1);'
                                }}
                                h1{{
                                    'color:#333;'
                                }}
                                p{{
                                    'margin:1em 0;'
                                    'line-height:1.5;'
                                }}
                            </style>
                        </head>
                        <body>
                            <div class='container'>
                                <h2>Please find the automated rsbcl report from {0}</h2>
                                <p>Here is the OneDrive public folder for rsbcl <a style="background-color: #4CAF50; border: none; color: white; padding: 5px 10px; text-align: center; text-decoration: none; display: inline-block; font-size: 16px; margin: 4px 2px; cursor: pointer; border-radius: 4px;" href='{1}'>RSBCL ONE DRIVE</a></p>
                                <p>Or you can directly download latest file by <a style="background-color: #4CAF50; border: none; color: white; padding: 5px 10px; text-align: center; text-decoration: none; display: inline-block; font-size: 16px; margin: 4px 2px; cursor: pointer; border-radius: 4px;" href='{2}'>Download</a>
                                <p><cite style='color:blue;'>(Please don't reply to this email as it is an automated message.Mail to mastwalrk@radico.co.in for any further queries)</cite></p>
                                <h3>Thank You !</h3>
                            </div>
                        </body>
                    </html>
                    '''.format(self.sheet_name_email,self.ONE_DRIVE_FOLDER,res)
            else:
                raise Exception('There has been some error while uploading report to one drive.')
            send_mail(
                to = self.to,
                from_email = self.from_email,
                password = self.password,
                subject=self.subject,
                cc = self.cc,
                body = body
            )
        except Exception as e:
            # driver.save_screenshot('error.png')
            # driver.close()
            raise e
        finally:
            driver.quit()
            