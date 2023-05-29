from excise_automation import ExciseAutomation
from send_mail import send_mail
import time
if __name__ == '__main__':
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            print(f'Attempt number: {attempt+1}')
            obj = ExciseAutomation()
            obj.run_driver()
        except Exception as e:
            print(e)
            if attempt == max_attempts - 1:
                body='''
                        <html>
                            <head>
                                <style>
                                    body{{
                                        font-family: Arial, sans-serif;
                                        background-color: #f2f2f2;
                                    }}
                                    .container{{
                                        width: 80%;
                                        margin: 0 auto;
                                        padding: 20px; 
                                        background-color: #fff;
                                        border-radius: 10px;   
                                        box-shadow: 0px 2px 5px rgba(0, 0, 0, 0.1);
                                    }}
                                    h1{{
                                        text-align:center;
                                        color:#333;
                                    }}
                                    p{{
                                        margin:1em 0;
                                        line-height:1.5;
                                    }}
                                    .error-msg {{
                                        color: #ff0000;
                                        font-weight: bold;
                                    }}
                                </style>
                            </head>
                                <body>
                                    <div class='container'>
                                        <p>We regret to inform you that an error occured while running the script.</p>
                                        <p class='error-msg'>Error message: {0}</p>
                                        <p>Thank you for your attention to this matter.</p>
                                        <p><cite style='color:blue;'>(Please don't reply to this email as it is an automated message.Mail to mastwalrk@radico.co.in or yogeshk@radico.co.in to resolve this issue)</cite></p>
                                        <h3>Thank You !</h3>
                                    </div>
                                </body>
                        </html>
                    '''.format(str(e))
                send_mail(obj.to,obj.from_email,obj.password,obj.cc,obj.subject,body,'error.png')
            else:
                time.sleep(5)
        else:
            break