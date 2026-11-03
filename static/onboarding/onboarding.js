(()=>{'use strict';
const ok=document.getElementById('rules-ok'),next=document.getElementById('continue'),panel=document.getElementById('rules-panel'),form=document.getElementById('onboarding-form');
ok.addEventListener('change',()=>next.disabled=!ok.checked);
next.addEventListener('click',()=>{if(!ok.checked)return;panel.style.display='none';form.style.display='block';form.scrollIntoView({behavior:'smooth'});resize();});
const canvas=document.getElementById('signature'),ctx=canvas.getContext('2d');let drawing=false,strokes=0;
function resize(){const rect=canvas.getBoundingClientRect();if(!rect.width)return;canvas.width=Math.round(rect.width*devicePixelRatio);canvas.height=Math.round(rect.height*devicePixelRatio);ctx.scale(devicePixelRatio,devicePixelRatio);ctx.fillStyle='#fff';ctx.fillRect(0,0,rect.width,rect.height);ctx.strokeStyle='#161616';ctx.lineWidth=2.7;ctx.lineCap='round';ctx.lineJoin='round';strokes=0;}
function position(e){const r=canvas.getBoundingClientRect();return {x:e.clientX-r.left,y:e.clientY-r.top};}
canvas.addEventListener('pointerdown',e=>{drawing=true;canvas.setPointerCapture(e.pointerId);const p=position(e);ctx.beginPath();ctx.moveTo(p.x,p.y);ctx.lineTo(p.x+.1,p.y+.1);ctx.stroke();strokes++;});
canvas.addEventListener('pointermove',e=>{if(!drawing)return;const p=position(e);ctx.lineTo(p.x,p.y);ctx.stroke();});
for(const n of ['pointerup','pointercancel'])canvas.addEventListener(n,()=>drawing=false);
document.getElementById('clear-signature').addEventListener('click',resize);
form.addEventListener('submit',e=>{if(!strokes){e.preventDefault();alert('Debes firmar el documento.');return;}document.getElementById('firma').value=canvas.toDataURL('image/png');});
})();
