from django.shortcuts import render
from django.template.loader import render_to_string
from django.http import JsonResponse, HttpResponse
from .scripts.Employee import Employee
import json


def SalaryCalculation(request):
    """Open the initial page"""
    context = {}
    return render(request, "Salary/index.html", context)


def DownloadFile(request):
    """
        Download Link when the user presses on the button.
        :returns a response that incldues the requested file.
    """

    filename = "Mazars_Gross_Net_Calculation.xlsx"  # Nmae of the files needs to be downloaded

    with open('Salary/scripts/Results/' + filename, 'rb') as f:
        response = HttpResponse(f.read(), content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename=' + filename
        response['Content-Type'] = 'application/vnd.ms-excel; charset=utf-16'
        return response


def renderTable(request):
    """
        Partial render
        get the user Inputs and run the calculations
        :returns jsonresponse that includes the html of the table only with the data that
                 will be written to the table
    """

    usrInputs = {}  # a dictionary to save the user inputs coming from the UI

    for k in request.GET:  # request.GET = {key:value}
        if request.GET[k].isdigit():  # check if the value can be a digit '1'
            usrInputs[k] = int(request.GET[k])
        elif request.GET[k] in ['false', 'true']:  # check if the 'false' is boolean or not;
            usrInputs[k] = json.loads(request.GET[k])
        else:  # for the json type object '{key:value}'
            usrInputs[k] = request.GET[k]

    """
        0: gross to net
        1: net to gross
    """

    if (usrInputs['calType'] == 0):  # check the operation flag
        gross_to_net = 1
        net_to_gross = 0
    else:
        gross_to_net = 0
        net_to_gross = 1

    # turning the json object into a nested dictionary inside usrInputs with the key salaries
    usrInputs['salaries'] = dict([(int(key), value) for key, value in json.loads(usrInputs['salaries']).items()])

    EmpObject = Employee(usrInputs['kidsCount'], usrInputs['socialStatus'], usrInputs['partnerStatus'],
                         usrInputs['empCost'],
                         usrInputs['empSocialShare'], gross_to_net, net_to_gross, usrInputs["salaries"])

    EmpObject.saveToFile()

    data = dict()
    data['html_form'] = render_to_string('Salary/table.html', {'result': EmpObject.data})

    return JsonResponse(data)


def fileReader(fileName):
    """:param parameters files name
       :return a dictionary of the title and the value of taxes mentioned in the file."""
    file = open('Salary/scripts/parameter/parameter' + fileName + ".txt", 'r')

    readings = file.read().splitlines()  # read and split o new lines
    data = {}
    for line in readings[1:-1]:  # ignore the first line and the last line => (\n)
        key, value = line.split(':')
        data[key] = value
    return data


def paremeters(request):
    """
        function that will respond to the paramteres button press
        :returns JsonResponse for the html code that corresponds for the modal and the data will be written to it
    """
    data = dict()
    year = request.GET["year"]  # get the year coming from the user press
    title = "Parameters For The Year " + year  # title for the modal
    parameters = fileReader(year)  # get the data inside the file related to the year

    # add the response html_form key that will have the html content with the data that will be passed to prameters.html
    data["html_form"] = render_to_string("Salary/paremeters.html", {'title': title, 'parameters': parameters})

    return JsonResponse(data)





