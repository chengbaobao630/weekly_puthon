#!/usr/bin/python
# -*- coding: UTF-8 -*-
import json
from weekly.mail import weekly_send

from django.http import HttpResponse


def send(request):
    if request.method == 'POST':
        received_json_data = json.loads(request.body)
        weekly_send.gen_send(received_json_data)
        return HttpResponse(json.dumps(received_json_data), "application/json")
