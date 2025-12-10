from django.shortcuts import render
from django.http import HttpResponse
from .forms import UploadForm
from .services import extract_all, generate_tn_ved_excel
import io

def process_document(request):
    if request.method == 'POST':
        form = UploadForm(request.POST, request.FILES)
        if form.is_valid():
            docx_file = request.FILES['file']
            url = form.cleaned_data['url']
            
            try:
                # Process the file
                abbreviations, units, tn_ved_codes = extract_all(docx_file)
                
                # Generate Excel
                wb = generate_tn_ved_excel(units, tn_ved_codes, url)
                
                # Save to buffer
                buffer = io.BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                
                # Create response
                response = HttpResponse(
                    buffer.getvalue(),
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                response['Content-Disposition'] = 'attachment; filename="tn_ved_processed.xlsx"'
                response.set_cookie('fileDownload', 'true', max_age=20)
                return response
                
            except Exception as e:
                form.add_error(None, f"Error processing file: {str(e)}")
    else:
        form = UploadForm()
    
    return render(request, 'processor/index.html', {'form': form})
