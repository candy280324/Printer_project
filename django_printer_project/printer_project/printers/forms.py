from django import forms

class GatePassForm(forms.Form):
    gate_pass_no   = forms.CharField(max_length=30,  label="Gate Pass No")
    associate_name = forms.CharField(max_length=50,  label="Associate Name")
    code           = forms.CharField(max_length=20,  label="Code")
    department     = forms.CharField(max_length=30,  label="Department")
    date           = forms.DateField(widget=forms.DateInput(attrs={"type": "date"}))
    out_time       = forms.TimeField(widget=forms.TimeInput(attrs={"type": "time"}))
    purpose        = forms.CharField(max_length=80,  label="Purpose of Going")
