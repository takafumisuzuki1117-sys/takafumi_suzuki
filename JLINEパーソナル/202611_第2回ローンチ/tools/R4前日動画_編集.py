import subprocess, glob
U='/root/.claude/uploads/d102eec0-e923-5501-b9f3-772d1da54f4c'
src={n: glob.glob(f'{U}/*IMG_7971_{n}.mov')[0] for n in (2,3,4)}
# (part, start, end, [(rel_start, text), ...])  rel_start = seconds from segment start
SEG=[
 (2,14.20,15.98,[(0,'こんばんは、\\N鈴木崇文です。')]),
 (2,24.10,28.55,[(0,'いよいよ明日、\\N夜7時半から'),(2.15,'「サッカーIQ爆上げ\\Nセミナー」です。')]),
 (2,33.45,40.25,[(0,'明日は、実際の試合映像を\\N3本使って、'),(3.55,'「なぜこのプレーに\\Nなったのか」を\\N一緒に考えていきます。')]),
 (3,4.55,9.50,[(0,'試合の見方が分かると、'),(2.25,'お子さんのサッカーの\\N見え方が変わってきます。')]),
 (3,13.15,16.25,[(0,'できれば、お子さんと一緒に\\Nぜひご参加ください。')]),
 (4,5.95,9.10,[(0,'とはいえ、夜7時半という\\N時間帯もあるので、')]),
 (4,9.45,14.10,[(0,'顔を出しての参加が\\N難しいご家庭も\\Nあるかなと思います。')]),
 (4,14.50,19.10,[(0,'その場合は、カメラオフで\\Nチャットだけでも\\N大丈夫ですので、'),(2.75,'ぜひご参加ください。')]),
 (4,25.05,29.45,[(0,'もし直前で都合が\\N悪くなってしまったら、'),(2.15,'LINEで教えて\\Nいただければ助かります。')]),
 (4,31.60,36.30,[(0,'それでは明日、\\NZoomでお会いしましょう！'),(2.15,'楽しみにしています。\\Nよろしくお願いします。')]),
]
def ts(t):
    h=int(t//3600); m=int(t%3600//60); s=t%60
    return f'{h}:{m:02d}:{s:05.2f}'
ass=['[Script Info]','ScriptType: v4.00+','PlayResX: 720','PlayResY: 1280','WrapStyle: 2','',
 '[V4+ Styles]','Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding',
 'Style: Sub,IPAPGothic,44,&H00FFFFFF,&H00FFFFFF,&H00000000,&H96000000,1,0,0,0,100,100,0,0,1,4,2,2,40,40,150,1',
 'Style: Tel,IPAPGothic,58,&H00FFFFFF,&H00FFFFFF,&H002A12C8,&HC0000000,1,0,0,0,100,100,0,0,3,14,0,8,40,40,170,1',
 '','[Events]','Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text']
vf=[]; af=[]; t=0
for i,(p,s,e,subs) in enumerate(SEG):
    d=e-s
    z = 1.0 if i%2==0 else 1.07
    cw=int(360/z)//2*2; ch=int(640/z)//2*2
    vf.append(f'[{(p-2)}:v]trim={s}:{e},setpts=PTS-STARTPTS,crop={cw}:{ch},scale=720:1280:flags=lanczos,setsar=1,fps=30[v{i}]')
    af.append(f'[{(p-2)}:a]atrim={s}:{e},asetpts=PTS-STARTPTS,afade=t=in:st=0:d=0.03,afade=t=out:st={d-0.06:.3f}:d=0.06[a{i}]')
    for k,(r,txt) in enumerate(subs):
        st=t+r; en=t+(subs[k+1][0] if k+1<len(subs) else d)
        ass.append(f'Dialogue: 0,{ts(st)},{ts(en-0.02)},Sub,,0,0,0,,{txt}')
    t+=d
total=t
ass.append(f'Dialogue: 1,{ts(total-2.0)},{ts(total)},Tel,,0,0,0,,{{\\fad(150,0)}}明日 19:30〜｜Zoom')
open('r4.ass','w',encoding='utf-8').write('\n'.join(ass)+'\n')
n=len(SEG)
fc=';'.join(vf+af)+';'+''.join(f'[v{i}][a{i}]' for i in range(n))+f'concat=n={n}:v=1:a=1[cv][ca];[cv]subtitles=r4.ass:fontsdir=/usr/share/fonts/opentype/ipafont-gothic[vo];[ca]loudnorm=I=-16:TP=-1.5:LRA=11,aresample=48000[ao]'
cmd=['ffmpeg','-v','error','-y','-i',src[2],'-i',src[3],'-i',src[4],'-filter_complex',fc,'-map','[vo]','-map','[ao]',
     '-c:v','libx264','-preset','slow','-crf','22','-profile:v','high','-pix_fmt','yuv420p','-c:a','aac','-b:a','128k','-ar','48000','-movflags','+faststart','R4_前日動画.mp4']
subprocess.run(cmd,check=True)
print('total',round(total,2))
