// Paste the deployed Apps Script web-app URL here. No login credentials or secrets.
window.REGISTRATION_FORM_URL = 'https://script.google.com/a/macros/diit.edu.bd/s/AKfycbz8G1aheVeGIxKzsxIMXl783fUMAG-7TZ_XW2pFvoprh6BYJYj_xIl59_B5vPDbFqWB9A/exec';

/* GitHub Pages cannot safely embed every Google Apps Script deployment.
   Open registration directly instead, avoiding the Google Drive iframe error. */
document.addEventListener('DOMContentLoaded', () => {
  const panel=document.querySelector('.formbox');
  const frame=document.getElementById('registrationFrame');
  if(frame)frame.remove();
  const setup=document.getElementById('setup');
  if(setup)setup.hidden=true;
  if(!panel)return;
  const box=document.createElement('div');box.className='registration-help';
  const mark=document.createElement('div');mark.className='registration-monogram';mark.textContent='✦';box.append(mark);
  const title=document.createElement('h3');title.textContent='One last memory together';box.append(title);
  const description=document.createElement('p');
  description.textContent='Complete the secure student form to receive your registration reference. After payment is verified, your QR entry ticket will be emailed.';
  box.append(description);
  const url=String(window.REGISTRATION_FORM_URL||'').trim();
  const valid=/^https:\/\/script\.google\.com\/(?:a\/macros\/diit\.edu\.bd\/s|macros\/s)\/[A-Za-z0-9_-]+\/exec$/.test(url);
  if(valid){
    const link=document.createElement('a');link.className='button launch-registration';link.href=url;
    link.target='_blank';link.rel='noopener noreferrer';link.textContent='Open Secure Registration ↗';box.append(link);
    const note=document.createElement('p');note.className='form-footnote';note.textContent='The registration form opens in a new tab. Please contact the coordinator if you have trouble accessing it.';box.append(note);
  }else{
    const note=document.createElement('p');note.className='registration-status';
    note.textContent='Online registration is being connected. Contact the coordinator for assistance.';box.append(note);
  }
  panel.append(box);
  const heroTitle=document.querySelector('.hero h1');
  if(heroTitle)heroTitle.innerHTML='The end of a chapter.<br><em>The start of everything.</em>';
});
