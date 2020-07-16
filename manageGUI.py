import os, sys
from functools import partial
import datetime
import PyPDF2
import tkinter as tk
import tkinter.filedialog as filedialog

class GUI():
  def __init__(self):
    # メインウィンドウ
    self._root = tk.Tk()
    self._root.title('PDFmerge')

    # Frame1
    self._frame1 = tk.Frame(self._root, padx=10, pady=10)
    self._frame1.grid(row=0, column=1, sticky=tk.E)
    # 「ファイル参照」ラベルの作成
    self._IFileLabel = tk.Label(self._frame1, text='ファイル選択>>', padx=5, pady=2)
    self._IFileLabel.pack(side=tk.LEFT)
    # 「ファイル参照」ボタンの作成
    self._IFileButton = tk.Button(self._frame1, text='選択', command=self._filedialog)
    self._IFileButton.pack(side=tk.LEFT)

    # Frame2
    self._frame2 = tk.Frame(self._root, padx=10, pady=10)
    self._frame2.grid(row=2, column=1, sticky=tk.E)
    self._headLabel = tk.Label(self._frame2, text='----------------------------------------PDFファイル一覧----------------------------------------')
    self._headLabel.pack(side=tk.TOP)
    # merge する PDF 一覧とそれらを操作するボタン一覧
    self._PDFLabels = []
    self._UpButtons = []
    self._DownButtons = []
    self._DelButtons = []
    self._subFrames = []
    for i in range(20):
      self._subFrames.append(tk.Frame(self._root))
      self._subFrames[-1].grid(row=3+i, column=1, sticky=tk.E)
      self._PDFLabels.append(tk.Label(self._subFrames[-1], padx=5, pady=2))
      self._PDFLabels[-1].pack(side=tk.LEFT)
      self._UpButtons.append(tk.Button(self._subFrames[-1], text='↑', command=partial(self._UpPDFLabel, i)))
      self._UpButtons[-1].pack(side=tk.LEFT)
      self._DownButtons.append(tk.Button(self._subFrames[-1], text='↓', command=partial(self._DownPDFLabel, i)))
      self._DownButtons[-1].pack(side=tk.LEFT)
      self._DelButtons.append(tk.Button(self._subFrames[-1], text='×', command=partial(self._DelPDFLabel, i)))
      self._DelButtons[-1].pack(side=tk.LEFT)
    
    # Frame3
    self._frame3 = tk.Frame(self._root, padx=10, pady=10)
    self._frame3.grid(row=23, column=1, sticky=tk.E)
    # 「新しく作成するPDFの名前」のラベル作成
    self._NewPDFLabel = tk.Label(self._frame3, text='新しく作成するPDFの名前>>', padx=5, pady=5)
    self._NewPDFLabel.pack(side=tk.LEFT)
    # 「新しく作成するPDFの名前」のテキストエリア作成
    self._entry3 = tk.StringVar()
    self._NewEntry = tk.Entry(self._frame3, textvariable=self._entry3, width=30)
    self._NewEntry.pack(side=tk.LEFT)
    # 「PDFをmergeする」のボタン作成
    self._NewPDFButton = tk.Button(self._frame3, text='mergeする', command=self._mergePDF)
    self._NewPDFButton.pack(fill='x', padx=30, side=tk.LEFT)

  # ファイル参照
  def _filedialog(self):
    fTyp = [('', 'pdf')]
    iFile = os.path.abspath(os.path.dirname(__file__))
    iFilePath = filedialog.askopenfilenames(filetype=fTyp, initialdir=iFile)
    for label, path in zip(self._PDFLabels, iFilePath):
      label['text'] = path

  # フォルダ参照
  def _dirdialog(self):
    iDir = os.path.abspath(os.path.dirname(__file__))
    iDirPath = filedialog.askdirectory(initialdir=iDir)
    return iDirPath

  # PDF合成
  def _mergePDF(self):
    # 合成元のPDF一覧
    merger = PyPDF2.PdfFileMerger()
    PDFPaths = [PDFLabel['text'] for PDFLabel in self._PDFLabels]
    for PDFPath in PDFPaths:
      if PDFPath:
        merger.append(PDFPath)
    # PDFの合成
    DirPath = self._dirdialog()
    NewPath = self._NewEntry.get()
    if NewPath:
      merger.write(DirPath+'/'+NewPath+'.pdf')
    else:
      # ファイル名が指定されていない場合，現在の日時をファイル名にする
      dt_now = datetime.datetime.now()
      merger.write(DirPath+'/'+str(dt_now).replace(':', '')+'.pdf')
    merger.close()
  
  # PDFLabels の操作
  # 指定した PDFLabel をひとつ上に持ってくる
  def _UpPDFLabel(self, index):
    if index > 0:
      if self._PDFLabels[index]['text']:
        self._PDFLabels[index]['text'], self._PDFLabels[index-1]['text'] = self._PDFLabels[index-1]['text'], self._PDFLabels[index]['text']
  # 指定した PDFLabel をひとつ下に持っていく
  def _DownPDFLabel(self, index):
    if index < 19:
      if self._PDFLabels[index+1]['text']:
        self._PDFLabels[index]['text'], self._PDFLabels[index+1]['text'] = self._PDFLabels[index+1]['text'], self._PDFLabels[index]['text']
  # 指定した PDFLabel を削除する
  def _DelPDFLabel(self, index):
    for i in range(index, 19):
      self._PDFLabels[i]['text'] = self._PDFLabels[i+1]['text']
    self._PDFLabels[19]['text'] = ''

  def show(self):
    self._root.mainloop()