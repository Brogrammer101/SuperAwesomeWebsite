import requests
import win32com.client as win32



def get_random_joke():
    url = "https://dad-jokes.p.rapidapi.com/random/joke"

    headers = {
	"X-RapidAPI-Key": "85ccfd1ff1msh4d21aa0b47a6b23p154261jsn1cf41446ae60",
	"X-RapidAPI-Host": "dad-jokes.p.rapidapi.com"
}

    response = requests.get(url, headers=headers)
    joke_data = response.json()

    setup = joke_data['body'][0]['setup']
    punchline = joke_data['body'][0]['punchline']

    return setup, punchline


def send_email(recipients):
    setup, punchline = get_random_joke()

    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = 'Joke of the Day'
    mail.HTMLBody = f"<p>[{setup}<br><br>{punchline}]</p>"
    mail.To = ";".join(recipients)
    mail.Send()

recipients = ['codeTesting578@outlook.com', 'codeTesting888@outlook.com']

send_email(recipients)
    


