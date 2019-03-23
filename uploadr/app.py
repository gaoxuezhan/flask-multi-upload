from flask import Flask, request, redirect, url_for, render_template
import os
import json
import glob
import zipfile
from uuid import uuid4

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    """Handle the upload of a file."""
    form = request.form

    # Create a unique "session ID" for this particular batch of uploads.
    upload_key = str(uuid4())

    # Is the upload using Ajax, or a direct POST by the form?
    is_ajax = False
    if form.get("__ajax", None) == "true":
        is_ajax = True

    # Target folder for these uploads.
    target = "uploadr/static/uploads/{}".format(upload_key)
    try:
        os.mkdir(target)
        os.mkdir(target + "/input")
        os.mkdir(target + "/output")
    except:
        if is_ajax:
            return ajax_response(False, "Couldn't create upload directory: {}".format(target))
        else:
            return "Couldn't create upload directory: {}".format(target)

    print("=== Form Data ===")
    for key, value in list(form.items()):
        print(key, "=>", value)

    for upload in request.files.getlist("file"):
        filename = upload.filename.rsplit("/")[0]
        destination = "/".join([target, filename])
        print("Accept incoming file:", filename)
        print("Save it to:", destination)
        upload.save(destination)

    # gaoxz
    print(destination)
    zf = zipfile.ZipFile(destination, "r")
    for fileM in zf.namelist():
        zf.extract(fileM, target + "/input")
    zf.close()

    # gaoxz2
    # path = target + "/input"  # 文件夹目录
    # files = os.listdir(path)  # 得到文件夹下的所有文件名称
    # s = []
    # for file in files:  # 遍历文件夹
    #     if not os.path.isdir(file):  # 判断是否是文件夹，不是文件夹才打开
    #         f = open(path + "/" + file);  # 打开文件
    #         iter_f = iter(f);  # 创建迭代器
    #         temp = ""
    #         for line in iter_f:  # 遍历文件，一行行遍历，读取文本
    #             temp = temp + line
    #         s.append(temp)  # 每个文件的文本存到list中

    input_path = target + "/input"  # 文件夹目录
    zip_path(input_path, target + "/output", "result.zip")

    if is_ajax:
        return ajax_response(True, upload_key)
    else:
        return redirect(url_for("upload_complete", uuid=upload_key))


@app.route("/files/<uuid>")
def upload_complete(uuid):
    """The location we send them to at the end of the upload."""

    # Get their files.
    root = "uploadr/static/uploads/{}".format(uuid)
    if not os.path.isdir(root):
        return "Error: UUID not found!"

    files = []
    for file in glob.glob("{}/*.*".format(root)):
        fname = file.split(os.sep)[-1]
        files.append(fname)

    return render_template("files.html", uuid=uuid, files=files, )


def ajax_response(status, msg):
    status_code = "ok" if status else "error"
    return json.dumps(dict(
        status=status_code,
        msg=msg,
    ))


def dfs_get_zip_file(input_path, result):
    files = os.listdir(input_path)
    for file in files:
        if os.path.isdir(input_path+'/' + file):
            dfs_get_zip_file(input_path + '/' + file, result)
        else:
            result.append(input_path + '/' + file)


def zip_path(input_path, output_path, output_name):
    f = zipfile.ZipFile(output_path+'/'+output_name, 'w', zipfile.ZIP_DEFLATED)
    filelists = []
    dfs_get_zip_file(input_path, filelists)
    for file in filelists:
        f.write(file, file.replace(input_path, ""))
    f.close()
