from django.shortcuts import render
from django.http import HttpResponse
from .utils import generate_mediation_docx

def index(request):
    return render(request, 'doc_gen/index.html')

def generate_docx(request):
    if request.method == "POST":
        # Capture data from the HTML form
        data = {
            'client_name': request.POST.get('client_name'),
            'branch_address': request.POST.get('branch_address'),
            'branch_address_corr': request.POST.get('branch_address_corr'),
            'telephone': request.POST.get('telephone'),
            'mobile': request.POST.get('mobile'),
            'email': request.POST.get('email'),
            
            'customer_name': request.POST.get('customer_name'),
            'address1': request.POST.get('address1'),
            'address_corr': request.POST.get('address_corr'),
            'op_telephone': request.POST.get('op_telephone'),
            'op_mobile': request.POST.get('op_mobile'),
            'op_email': request.POST.get('op_email'),
        }
        
        # 1. Generate Word file in memory (RAM)
        buffer = generate_mediation_docx(data)
        
        # 2. Upload to Google Drive
        buffer.seek(0) # Move to start of file for reading
        file_name = f"Mediation_{data.get('client_name')}.docx"
        

        # 3. Provide the file for local download
        buffer.seek(0) # Reset again for the download response
        response = HttpResponse(
            buffer.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        response['Content-Disposition'] = f'attachment; filename="{file_name}"'
        return response

    return render(request,'doc_gen/index.html')