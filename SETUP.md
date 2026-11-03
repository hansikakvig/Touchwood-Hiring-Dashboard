# Touchwood Hiring Dashboard – Setup

## 1. Project banao
1. https://script.google.com par jao, us Google account se jiske paas application sheets ka access hai (social@brandandbeeyond.com sabse sahi hai, sheets usi ki hain).
2. **New project** banao, naam do "Touchwood Hiring".
3. `Code.gs` ka poora content editor ki `Code.gs` file mein paste karo.
4. **+ → HTML** se nayi file banao, naam `Index` rakho (extension `.html` apne aap lagega). Isme `Index.html` ka poora content paste karo.

## 2. Settings badlo (Code.gs ke upar CONFIG mein)
- `FIRST_ADMIN_PASSWORD`: apna strong password rakho.
- `LOGO_FILE_ID`: Touchwood World logo (PNG/JPG) Drive par upload karo aur uska link yahan paste karo. Baad mein Setup page se bhi badal sakte ho.

## 3. Setup chalao
1. Function dropdown mein `setup` chuno, **Run** dabao.
2. Permissions maangega (Sheets, Drive, triggers). **Allow** karo.
3. Isse ye banega:
   - Database sheet: "Touchwood Hiring Database"
   - Resume folder: "Touchwood Hiring Resumes"
   - Pehla admin user
   - 4 openings jo aapki application sheets se jude hain
   - 10 minute wala auto-import trigger

## 4. Deploy karo
1. **Deploy → New deployment → Web app**.
2. **Execute as:** Me.
3. **Who has access:** Anyone.
4. Deploy karke jo web app link mile, wahi dashboard ka link hai.

## 5. Login aur users
- **Pehla login:** Username `admin` aur CONFIG wala password. Login karke turant **Setup → Change password** karo.
- **Naye user:** Setup → Users who can sign in → Add user. Role "HR user" ya "Admin".
  - HR user: openings, candidates aur interviews manage kar sakta hai.
  - Admin: upar ka sab, plus users aur logo.
- **Password bhool gaye:** Admin us user ko Edit karke naya password set kar de.
- **Security:**
  - 5 galat password par account 15 minute ke liye lock hota hai.
  - 30 minute kuch na karo to sign out ho jata hai.
  - Passwords hashed (encrypted) save hote hain.

## Code badalne ke baad
Deploy → Manage deployments → Edit → Version: **New version** → Deploy. Link wahi rehta hai.
