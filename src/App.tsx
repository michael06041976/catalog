import React, { useState, useEffect } from 'react';
import { useAuthState } from 'react-firebase-hooks/auth';
import { auth, db } from './firebase';
import { doc, onSnapshot } from 'firebase/firestore';
import { updatePassword, EmailAuthProvider, reauthenticateWithCredential } from 'firebase/auth';
import Auth from './Auth';
import MainApp from './MainApp';
import { LogOut, Key } from 'lucide-react';

export default function App() {
  const [user, loading] = useAuthState(auth);
  const [role, setRole] = useState<string | null>(null);
  const [menus, setMenus] = useState<string[] | null>(null);
  const [status, setStatus] = useState<string | null>(null);
  const [showPasswordChange, setShowPasswordChange] = useState(false);
  const [currentPassword, setCurrentPassword] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [passError, setPassError] = useState('');

  useEffect(() => {
    if (user) {
      const unsub = onSnapshot(doc(db, 'users', user.uid), async (d) => {
        if (d.exists()) {
          const data = d.data();
          setRole(data.role);
          setMenus(data.menus || null);
          if (data.role === 'admin' && data.status === 'inactive') {
            const { updateDoc } = await import('firebase/firestore');
            updateDoc(doc(db, 'users', user.uid), { status: 'approved' }).catch(console.error);
            setStatus('approved');
          } else if (user.email === 'mike.allouche@outlook.com' && (data.role !== 'admin' || data.status === 'inactive')) {
            const { updateDoc } = await import('firebase/firestore');
            updateDoc(doc(db, 'users', user.uid), { status: 'approved', role: 'admin' }).catch(console.error);
            setStatus('approved');
            setRole('admin');
          } else {
            setStatus(data.status || 'approved');
          }
        } else {
          // Check for pre-added whitelist by Admin
          if (user.email) {
            try {
              const idByEmail = user.email.replace(/[@.]/g, '_');
              const { getDoc, setDoc, deleteDoc } = await import('firebase/firestore');
              const emailDoc = await getDoc(doc(db, 'users', idByEmail));
              
              if (emailDoc.exists()) {
                const data = emailDoc.data();
                setRole(data.role);
                setMenus(data.menus || null);
                setStatus(data.status || 'approved');
                
                // Migrate to UID-based doc
                await setDoc(doc(db, 'users', user.uid), data);
                await deleteDoc(doc(db, 'users', idByEmail));
                return;
              }
            } catch (err: any) {
               if (err.code !== 'permission-denied') {
                 console.error("Migration error:", err);
               }
            }
          }

          // If no doc and the user is the bootstrapped admin, give admin role in UI context
          if (user.email === 'm.allouche@gmail.com' || user.email === 'mike.allouche@outlook.com') {
             setRole('admin');
             setMenus(['catalog', 'add', 'company', 'admin']);
             setStatus('approved');
             
             // Optionally ensure they have a document
             try {
                const { setDoc, serverTimestamp } = await import('firebase/firestore');
                await setDoc(doc(db, 'users', user.uid), {
                  email: user.email,
                  role: 'admin',
                  status: 'approved',
                  createdAt: serverTimestamp(),
                  menus: ['catalog', 'add', 'company', 'admin']
                });
             } catch (e) {}
          } else {
             // Create a document for standard users who logged in via Google but have no document
             try {
                const { setDoc, serverTimestamp } = await import('firebase/firestore');
                await setDoc(doc(db, 'users', user.uid), {
                  email: user.email,
                  role: 'viewer',
                  status: 'pending',
                  createdAt: serverTimestamp(),
                  menus: ['catalog']
                });
             } catch (err: any) {
                console.error("Could not create pending user", err);
                setRole('viewer');
                setStatus('pending');
             }
          }
        }
      });
      return () => unsub();
    } else {
      setRole(null);
      setMenus(null);
      setStatus(null);
    }
  }, [user]);

  const handleChangePassword = async (e: React.FormEvent) => {
     e.preventDefault();
     if (!user) return;
     try {
       const isPasswordProvider = user.providerData.some(p => p.providerId === 'password');
       if (!isPasswordProvider) {
         setPassError('חשבון זה מחובר באמצעות ספק חיצוני (כגון Google) ואין לו סיסמא לשנות במערכת זו.');
         return;
       }
       
       if (!currentPassword) {
         setPassError('יש להזין סיסמא נוכחית.');
         return;
       }

       if (user.email) {
          const cred = EmailAuthProvider.credential(user.email, currentPassword);
          await reauthenticateWithCredential(user, cred);
       }

       await updatePassword(user, newPassword);
       alert('הסיסמא שונתה בהצלחה');
       setShowPasswordChange(false);
       setCurrentPassword('');
       setNewPassword('');
       setPassError('');
     } catch (err: any) {
       console.error(err);
       if (err.code === 'auth/wrong-password' || err.code === 'auth/invalid-credential') {
          setPassError('הסיסמא הנוכחית שגויה.');
       } else if (err.code === 'auth/requires-recent-login') {
          setPassError('יש להתנתק ולהתחבר מחדש כדי לשנות סיסמא.');
       } else if (err.code === 'auth/weak-password') {
          setPassError('סיסמא חלשה. לפחות 6 תווים.');
       } else {
          setPassError(err.message);
       }
     }
  };

  if (loading) {
    return <div className="min-h-screen flex items-center justify-center font-sans">טוען...</div>;
  }

  if (!user) {
    return <Auth />;
  }

  if (status === 'pending') {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center font-sans bg-slate-50 text-slate-800 p-4" dir="rtl">
        <h2 className="text-2xl font-bold mb-2 text-indigo-900">חשבונך ממתין לאישור מנהל</h2>
        <p className="text-slate-600 mb-6 text-center max-w-md">הרשמתך נקלטה בהצלחה, אך נדרש אישור מנהל מערכת טרם התחברות למערכת.</p>
        <button onClick={() => auth.signOut()} className="px-6 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-md font-medium shadow-sm">התנתק</button>
      </div>
    );
  }

  if (status === 'inactive') {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center font-sans bg-slate-50 text-slate-800 p-4" dir="rtl">
        <h2 className="text-2xl font-bold mb-2 text-red-700">חשבונך הושעה</h2>
        <p className="text-slate-600 mb-6 text-center max-w-md">הגישה של חשבונך למערכת נחסמה על ידי מנהל. אנא צור קשר עם ההנהלה.</p>
        <button onClick={() => auth.signOut()} className="px-6 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-md font-medium shadow-sm">התנתק</button>
      </div>
    );
  }

  return (
    <>
      <div className="fixed bottom-4 left-4 z-50 flex flex-col gap-2">
        <button 
          onClick={() => setShowPasswordChange(true)} 
          className="bg-white border border-slate-200 text-slate-700 p-3 rounded-full hover:bg-slate-50 shadow-lg flex items-center justify-center text-sm w-12 h-12"
          title="שנה סיסמא"
        >
          <Key size={18} />
        </button>
        <button 
          onClick={() => auth.signOut()} 
          className="bg-slate-800 text-white p-3 rounded-full hover:bg-slate-700 shadow-lg flex items-center justify-center text-sm w-12 h-12"
          title="התנתק"
        >
          <LogOut size={18} />
        </button>
      </div>

      {showPasswordChange && (
         <div className="fixed inset-0 bg-slate-900/60 z-[1000] flex items-center justify-center p-4" dir="rtl">
            <div className="bg-white rounded-xl shadow-xl w-full max-w-sm p-6 relative">
               <h3 className="text-lg font-bold mb-4">שינוי סיסמא</h3>
               <form onSubmit={handleChangePassword} className="space-y-4">
                 {user?.providerData.some(p => p.providerId === 'password') && (
                   <div>
                     <label className="block text-sm mb-1 text-slate-700">סיסמא נוכחית</label>
                     <input 
                       type="password" 
                       required 
                       value={currentPassword}
                       onChange={e => setCurrentPassword(e.target.value)}
                       className="w-full border border-slate-200 rounded-md px-3 py-2 text-sm outline-none focus:border-indigo-500"
                     />
                   </div>
                 )}
                 <div>
                   <label className="block text-sm mb-1 text-slate-700">סיסמא חדשה</label>
                   <input 
                     type="password" 
                     required 
                     minLength={6}
                     value={newPassword}
                     onChange={e => setNewPassword(e.target.value)}
                     className="w-full border border-slate-200 rounded-md px-3 py-2 text-sm outline-none focus:border-indigo-500"
                   />
                 </div>
                 {passError && <p className="text-red-500 text-xs">{passError}</p>}
                 <div className="flex gap-2 justify-end pt-2">
                   <button type="button" onClick={() => setShowPasswordChange(false)} className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-sm font-medium rounded-md">ביטול</button>
                   <button type="submit" className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium rounded-md">עדכן סיסמא</button>
                 </div>
               </form>
            </div>
         </div>
      )}

      <MainApp userRole={role || 'viewer'} userMenus={menus} />
    </>
  );
}
