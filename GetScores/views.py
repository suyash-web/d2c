from django.shortcuts import render
from django.core.files.storage import FileSystemStorage
from django.core.files.uploadedfile import InMemoryUploadedFile
from django.contrib import messages
from pathlib import Path
from GetScores.tasks import process_file
import multiprocessing

# Create your views here.
def index(request):
    return render(request, "index.html")
    # return HttpResponseRedirect(request, "/success/")

def upload(request):
    BASE_DIR = Path(__file__).resolve().parent.parent
    context = {}
    if request.method == 'POST':
        print("post method")
        uploaded_file: InMemoryUploadedFile = request.FILES['document']
        fs = FileSystemStorage()
        name = fs.save(uploaded_file.name, uploaded_file)
        context['url'] = fs.url(name)
        process = multiprocessing.Process(target=process_file, args=(uploaded_file, BASE_DIR / "storage" / uploaded_file.name, uploaded_file.name))
        process.start()
        messages.warning(request, "Processing...")
        # messages.warning(request, "Processing...")
        # process_file(uploaded_file, BASE_DIR / "storage", uploaded_file.name)
    return render(request, 'index.html', context=context)
