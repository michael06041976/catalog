import React, { useState } from 'react';
import { useSignInWithGoogle, useSignInWithEmailAndPassword, useCreateUserWithEmailAndPassword } from 'react-firebase-hooks/auth';
import { setDoc, doc, serverTimestamp } from 'firebase/firestore';
import { auth, db } from './firebase';

export default function Auth() {
  const [signInWithGoogle, gUser, gLoading, gError] = useSignInWithGoogle(auth);
  const [signInWithEmailAndPassword, eUser, eLoading, eError] = useSignInWithEmailAndPassword(auth);
  const [createUserWithEmailAndPassword, cUser, cLoading, cError] = useCreateUserWithEmailAndPassword(auth);

  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [role, setRole] = useState('viewer');
  const [isRegistering, setIsRegistering] = useState(false);

  const handleEmailAuth = async (e: React.FormEvent) => {
    e.preventDefault();
    if (isRegistering) {
      const res = await createUserWithEmailAndPassword(email, password);
      if (res && res.user) {
        try {
          await setDoc(doc(db, 'users', res.user.uid), {
            email: res.user.email,
            role: role,
            status: 'pending',
            createdAt: serverTimestamp(),
            menus: role === 'viewer' ? ['catalog'] : (role === 'editor' ? ['catalog', 'add'] : ['catalog', 'add', 'company', 'admin'])
          });
        } catch (err) {
          console.error("Error setting user doc:", err);
        }
      }
    } else {
      signInWithEmailAndPassword(email, password);
    }
  };

  const loading = gLoading || eLoading || cLoading;
  const error = gError || eError || cError;

  return (
    <div className="min-h-screen bg-slate-100 flex items-center justify-center font-sans p-4" dir="rtl">
      <div className="bg-white p-6 md:p-8 rounded-xl shadow-sm border border-slate-200 max-w-sm w-full text-center">
        <div className="w-16 h-16 bg-rose-100 text-rose-600 rounded-full flex items-center justify-center mx-auto mb-4 text-3xl">
          ◈
        </div>
        <h2 className="text-2xl font-bold text-slate-800 mb-2">ברוך הבא למערכת</h2>
        <p className="text-slate-500 text-sm mb-6">התחבר כדי לנהל את הקטלוג ומשתמשי המערכת</p>
        
        <form onSubmit={handleEmailAuth} className="space-y-3 mb-6 relative">
          <input 
            type="email" 
            placeholder="אימייל" 
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            className="w-full px-3 py-2 border border-slate-200 rounded-md text-[14px] outline-none focus:border-indigo-400"
            required
            dir="ltr"
          />
          <input 
            type="password" 
            placeholder="סיסמא" 
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            className="w-full px-3 py-2 border border-slate-200 rounded-md text-[14px] outline-none focus:border-indigo-400"
            required
            minLength={6}
            dir="ltr"
          />
          
          {/* Role selection removed, defaults to viewer */}

          <button 
            type="submit"
            className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-medium py-2.5 px-4 rounded-md transition-colors"
            disabled={loading}
          >
            {loading ? 'טוען...' : (isRegistering ? 'צור חשבון חדש' : 'התחבר עם אימייל')}
          </button>
        </form>

        <div className="text-sm text-slate-500 flex items-center justify-center gap-2 mb-6">
          <span className="w-full h-px bg-slate-200"></span>
          <span>או</span>
          <span className="w-full h-px bg-slate-200"></span>
        </div>

        <button 
          onClick={() => signInWithGoogle()}
          className="w-full flex items-center justify-center gap-3 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 font-medium py-2.5 px-4 rounded-md transition-colors"
          disabled={loading}
        >
          <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" className="w-5 h-5" alt="Google" />
          {loading ? 'מתחבר...' : 'התחברות עם Google'}
        </button>

        <div className="mt-6 text-sm text-slate-600">
          {isRegistering ? 'יש לך כבר חשבון? ' : 'אין לך חשבון? '}
          <button 
            onClick={() => setIsRegistering(!isRegistering)}
            className="text-indigo-600 hover:text-indigo-800 font-medium"
            type="button"
          >
            {isRegistering ? 'התחבר' : 'הירשם כאן'}
          </button>
        </div>

        {error && <p className="text-red-500 text-sm mt-4 bg-red-50 p-2 rounded-md">{error.message}</p>}
      </div>
    </div>
  );
}
