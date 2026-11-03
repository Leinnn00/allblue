# ヒーロー動画づくり: 必要なもの = ffmpeg (PATHに通す) と Python
# 使い方: python tools/make_hero.py → photos/videos/hero_*_draft.mp4 ができる
import subprocess,sys
import os
V=os.path.join(os.path.dirname(os.path.abspath(__file__)),"..","photos","videos")+os.sep
# ★ここを編集: (ファイル名, 開始秒) の順番がそのままつなぐ順番
clips=[("drone.MOV",1),("meetball.MOV",12.5),("meetball.MOV",19),("squba.MP4",2),("jinbe.MP4",20),("shark.MP4",20),("meetball.MOV",54)]
D=3.5   # 1カットの長さ(秒)
F=0.5   # つなぎ目のフェード(秒)
def build(mode,out):
    args=["ffmpeg","-v","error","-y"]
    for f,s in clips: args+=["-ss",str(s),"-t",str(D),"-i",V+f]
    parts=[]
    for i in range(len(clips)):
        if mode=="pc": vf="crop=720:405,scale=1280:720:flags=lanczos"
        else: vf="scale=720:1280"
        parts.append(f"[{i}:v]{vf},fps=30,setsar=1,format=yuv420p[v{i}]")
    last="v0"; off=D-F
    for i in range(1,len(clips)):
        parts.append(f"[{last}][v{i}]xfade=transition=fade:duration={F}:offset={off}[x{i}]"); last=f"x{i}"; off+=D-F
    parts.append(f"[{last}]fade=t=in:st=0:d=0.5[o]")
    args+=["-filter_complex",";".join(parts),"-map","[o]","-an","-c:v","libx264","-crf","27","-preset","slow","-movflags","+faststart",out]
    subprocess.run(args,check=True)
build("pc",V+"hero_pc_draft.mp4"); build("sp",V+"hero_sp_draft.mp4")
