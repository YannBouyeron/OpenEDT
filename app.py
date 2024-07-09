import os
from bottle import*
from etab import Etab
from edt import testx
import zipfile
import random

app = Bottle()

@app.route('/static/<filepath:path>')
def send_static(filepath):

    return static_file(filepath, root='./static/')

@app.route('/server_static/<filepath:path>')
def server_static(filepath):

    return static_file(filepath, root='./etab/')

@app.route('/')
def index():

    # on supprime l’eventuel cookie pour pouvoir generer un nouvel edt

    response.set_cookie("edt", '' ,secret='hgcghgfygfytfftyf', expires=0)

    return template("index.html")

@app.post('/analyse')
def analyse():

    if request.forms.get("creat"):

        return createdt()

    upload  = request.files.get('upload')

    # on verifie l’extension
    name, ext = os.path.splitext(upload.filename)
    if ext not in ('.xlsx'):
        return 'File extension not allowed.'

    # on choisit un nom (nombre) au hasard
    n = random.choice(range(1000000, 2000000))

    # on enregistre le xlsx dans etab/
    upload.save("etab/{0}".format(n))

    try:

        e = Etab("etab/{0}".format(n))

    except:
        # par securité on ne conserve pas le fichier uploadé sur le serveur si il est défecteux
        os.remove("etab/{0}".format(n))

        return "<p>Votre fichier etab.xlsx n’est pas conforme !</p>"

    else:

        r , h = e.analyse()

        os.remove("etab/{0}".format(n))

        return h

@app.post('/createdt')
def createdt():

    # on essai de recuperer le cookies , si il est présent c’est que la page a ete rechargee et il n’est pas necessaire de régénerer un new edt

    n = request.get_cookie("edt", secret='hgcghgfygfytfftyf')

    if  n:

        # on read l’edt au format html dans etab
        with open("etab/{0}.html".format(n), "r") as h:
            ht = h.read()
            h.close()

        excel = """
    	    <form action="/server_static/{0}_EDT.zip" method="get" accept-charset="ISO-8859-1">
    	    <button type="submit" name="" value="" class = "">Télécharger votre EDT.xlsx</button>
    	    </form>
           </br>
           <form action="/" method="get" accept-charset="ISO-8859-1">
    	    <button type="submit" name="" value="" class = "">Créer un nouvel EDT</button>
    	    </form>
	    """.format(n)

        return template("edt.html", excel=excel, ht=ht)
	    
    # si le cookie n’est pas présent c’est que la route est appelée depuis l’index alors on upload et on genere un edt

    upload  = request.files.get('upload')

    # on verifie l’extension
    name, ext = os.path.splitext(upload.filename)
    if ext not in ('.xlsx'):
        return 'File extension not allowed.'

    # on choisit un nom (nombre) au hasard
    n = random.choice(range(1000000, 2000000))

    # on enregistre le xlsx dans etab/
    upload.save("etab/{0}".format(n))

    # on realise l’emploi du temps avec la fonction testx de edt.py
    try:

        x = testx("etab/{0}".format(n), n)

    except:

        # par securité on ne conserve pas le fichier uploadé sur le serveur si il est défecteux
        os.remove("etab/{0}".format(n))

        return "<p>Votre fichier etab.xlsx n’est pas conforme !</p>"

    # on save le df des seances de l’etablissement retourné complété par testx, avec les valeur dans regroup
    df = x[0]
    df.to_excel("etab/{0}.xlsx".format(n), index=False)
	
    # on réimporte ce meme fichier dans un objet Etab de etab.py
    e = Etab("etab/{0}.xlsx".format(n)) 

    # on le re save , mais cette fois avec la methode save de l’objet Etab, ce qui permet de remplacer les valeurs de regroup par des formules
    e.save()

    # on zip le fichier pour pouvoir conserver les fonctions excel des regroupements
    my_zipfile = zipfile.ZipFile("etab/{0}_EDT.zip".format(n), mode='w', compression=zipfile.ZIP_DEFLATED)
    my_zipfile.write("etab/{0}.xlsx".format(n))
    my_zipfile.write("etab/{0}_EDT.xlsx".format(n))
    my_zipfile.close()

    # on save l’edt au format html dans etab
    with open("etab/{0}.html".format(n), "w") as h:
        h.write(x[3])
        h.close()

    # on envoie un cookies pour conserver le random name en cas de rechargement de la page
    response.set_cookie("edt", n, secret='hgcghgfygfytfftyf', max_age=60000)

    # bouton pour download le xlsx
    excel = """
    	<form action="/server_static/{0}_EDT.zip" method="get" accept-charset="ISO-8859-1">
    	<button type="submit" name="" value="" class = "">Télécharger votre EDT.xlsx</button>
    	</form>
       </br>
       <form action="/" method="get" accept-charset="ISO-8859-1">
    	<button type="submit" name="" value="" class = "">Créer un nouvel EDT</button>
    	</form>
	""".format(n)

    return template("edt.html", excel=excel, ht=x[3])

@app.get('/createtab')
def createtab():

    return template("createtab.html")

@app.post('/createtab')
def createtab():

    nom = request.forms.get("nom")

    six = request.forms.get('Sixième')
    cin = request.forms.get('Cinquième')
    qua = request.forms.get('Quatrième')
    tro = request.forms.get('Troisième')
    sec = request.forms.get('Seconde')
    pre = request.forms.get('Première')
    ter = request.forms.get('Terminale')

    sam = request.forms.get('samedi')
    mer = request.forms.get('mercredi')
    m1 = request.forms.get('m1')
    a3 = request.forms.get('a3')
    a4 = request.forms.get('a4')

    d = {"sixième":int(six), "cinquième":int(cin), "quatrième":int(qua), "troisième":int(tro), "seconde":int(sec), "première":int(pre), "terminale":int(ter)}

    if sam == "1":
        d["samedi"]=True

    if mer == "1":
        d["mercredi"]=True

    if m1 == "1":
        d["m1"]=True

    if a3 == "1":
        d["a3"]=True

    if a4 == "1":
        d["a4"]=True

    e = Etab("etab/" + nom, **d)

    e.save()

    # on zip le fichier pour pouvoir conserver les fonctions excel des regroupements

    my_zipfile = zipfile.ZipFile("etab/{0}.zip".format(nom), mode='w', compression=zipfile.ZIP_DEFLATED)

    f = "etab/"+nom+".xlsx"

    my_zipfile.write(f)
    my_zipfile.close()


    # redirection automatique pour download etab.xlsx

    response.status = 303
    response.set_header('Location', '/server_static/{0}.zip'.format(nom))

app.run(host='127.0.0.1', port=27200, reload=True, debug=True)
