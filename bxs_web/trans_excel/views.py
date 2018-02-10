# from django.http import JsonResponse
from django.shortcuts import HttpResponse
from django.shortcuts import render
from bxs_web.settings import BASE_DIR
from django.views.decorators.csrf import csrf_exempt
import os

from .deal_forms import trans

# Create your views here.

@csrf_exempt
def upload(request):
    if request.method == 'POST':
        up_file = request.FILES.get('file', None)
        if not up_file or up_file.size < 1:
            # return HttpResponse("请选择文件", content_type="application/json")
            return HttpResponse("请选择文件")
        if '.xlsx' != up_file.name[up_file.name.rfind('.'):]:
            # return ("请将保存为 xlsx 后缀， 打开 Excel 保存为 2007 版本或更高版本")
            return HttpResponse ("请将保存为 xlsx 后缀， 打开 Excel 保存为 2007 版本或更高版本")
        file_path = os.path.join(BASE_DIR,'static', '1.xlsx')
        #删除所有 xlsx 文件
        items = os.listdir(os.path.join(BASE_DIR, 'static'))
        for i in items:
           if i.endswith('.xlsx'):
               os.remove(os.path.join(BASE_DIR, 'static', i))
        # if os.path.exists(file_path):
        #     os.remove(file_path)
        with open(file_path, 'wb+') as fp:
            for chunk in up_file.chunks():  # 分块写入文件
                fp.write(chunk)
        #调用处理程序
        out_file_paths= trans.main(file_path, os.path.join(BASE_DIR,'static'))
        return render(request, 'download.html', {'down_list': out_file_paths})
    else:
        return render(request, 'trans.html')