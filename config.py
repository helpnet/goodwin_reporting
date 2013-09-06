import datetime
import webbrowser

ticket_hold = []
now = datetime.datetime.now()
data_path = "Data/data " + str(now)[:-10] + ".csv"
report_path = "Data/report " + str(now)[:-10] + ".csv"
firefox = webbrowser.get('firefox')
uname = "admin"
url_login = "http://stream.goodwin.drexel.edu/learningtech/wp-login.php"
url_ticket_home = "http://stream.goodwin.drexel.edu/learningtech/ticketsystem/wp-admin/edit.php?post_type=ticket&ticket_status=all"
url_tickets = "http://stream.goodwin.drexel.edu/learningtech/ticketsystem/wp-admin/edit.php?post_type=ticket&ticket_status=all&paged="
soe_course = ['EDUC', 'EDHE', 'EDGI', 'EHRD', 'EDEX', 'EDD', 'EDLT']
sotaps_course = ['SMT', 'HRM', 'PROJ', 'REAL', 'CMGT', 'CULA', 'ET', 'CST', 'CT', 'MEP', 'HRD', 'CRTV', 'ORGB', 'GSTD', 'HSM'] 
