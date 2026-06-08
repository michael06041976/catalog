import React, { useState, useEffect } from 'react';
import { collection, onSnapshot, doc, setDoc, deleteDoc, serverTimestamp, updateDoc } from 'firebase/firestore';
import { sendPasswordResetEmail } from 'firebase/auth';
import { db, auth } from './firebase';
import { Shield, UserPlus, Trash2, ShieldAlert, Building, Edit, Key, X, Check } from 'lucide-react';

interface AdminProps {
  onNavigateToCompany: () => void;
}

const MENU_OPTIONS = [
  { id: 'catalog', label: 'קטלוג (צפייה)' },
  { id: 'add', label: 'הוספת/עריכת פריטים' },
  { id: 'quote', label: 'הצעת מחיר' },
  { id: 'company', label: 'הגדרות חברה' },
  { id: 'pricelists', label: 'ניהול מחירונים' },
  { id: 'admin', label: 'ניהול מערכת' }
];

export default function Admin({ onNavigateToCompany }: AdminProps) {
  const [users, setUsers] = useState<any[]>([]);
  const [newUserEmail, setNewUserEmail] = useState('');
  const [newUserRole, setNewUserRole] = useState('viewer');
  const [newUserMenus, setNewUserMenus] = useState<string[]>(['catalog']);

  const [newUserPhone, setNewUserPhone] = useState('');
  const [newUserCompany, setNewUserCompany] = useState('');

  const [editingUserId, setEditingUserId] = useState<string | null>(null);
  const [editRole, setEditRole] = useState('viewer');
  const [editMenus, setEditMenus] = useState<string[]>([]);
  const [editPhone, setEditPhone] = useState('');
  const [editCompany, setEditCompany] = useState('');

  useEffect(() => {
    const unsub = onSnapshot(collection(db, 'users'), (snapshot) => {
      setUsers(snapshot.docs.map(d => ({ id: d.id, ...d.data() })));
    }, (error) => {
      console.error(error);
    });
    return () => unsub();
  }, []);

  const handleAddUser = async () => {
    if (!newUserEmail) return;
    try {
      const id = newUserEmail.replace(/[@.]/g, '_');
      await setDoc(doc(db, 'users', id), {
        email: newUserEmail,
        role: newUserRole,
        menus: newUserMenus,
        phone: newUserPhone,
        company: newUserCompany,
        status: 'approved',
        createdAt: serverTimestamp()
      });
      setNewUserEmail('');
      setNewUserPhone('');
      setNewUserCompany('');
      setNewUserMenus(['catalog']);
    } catch (err: any) {
      alert("שגיאה בהוספת משתמש: " + err.message);
    }
  };

  const [userToDelete, setUserToDelete] = useState<string | null>(null);

  const handleDelUser = async (id: string) => {
    try {
      await deleteDoc(doc(db, 'users', id));
      setUserToDelete(null);
    } catch (err: any) {
      alert("שגיאה במחיקת משתמש: " + err.message);
    }
  };

  const handleResetPassword = async (email: string) => {
    try {
      await sendPasswordResetEmail(auth, email);
      alert(`מטעמי אבטחה של גוגל, לא ניתן להגדיר סיסמא ידנית עבור משתמשים אחרים.\n\nנשלח בהצלחה קישור לאיפוס סיסמא לכתובת: ${email}\nהמשתמש יוכל להגדיר סיסמא חדשה דרך הקישור במייל.`);
    } catch (err: any) {
      if (err.code === 'auth/user-not-found' || err.code === 'auth/invalid-email') {
         alert(`שגיאה: המשתמש עדיין לא ביצע הרשמה ראשונית למערכת עם המייל ${email}, ולכן לא ניתן לשלוח לו איפוס סיסמא.`);
      } else {
         alert("שגיאה בשליחת קישור איפוס: " + err.message);
      }
    }
  };

  const startEditUser = (u: any) => {
    setEditingUserId(u.id);
    setEditRole(u.role || 'viewer');
    setEditMenus(u.menus || []);
    setEditPhone(u.phone || '');
    setEditCompany(u.company || '');
  };

  const saveEditUser = async (id: string, newStatus?: string) => {
    try {
      await updateDoc(doc(db, 'users', id), {
        role: editRole,
        menus: editMenus,
        phone: editPhone,
        company: editCompany,
        ...(newStatus && { status: newStatus })
      });
      if (!newStatus) setEditingUserId(null);
    } catch (err: any) {
      alert("שגיאה בעדכון משתמש: " + err.message);
    }
  };

  const toggleMenu = (menuId: string, currentMenus: string[], setter: (m: string[]) => void) => {
    if (currentMenus.includes(menuId)) {
      setter(currentMenus.filter(m => m !== menuId));
    } else {
      setter([...currentMenus, menuId]);
    }
  };

  const changeUserStatus = async (id: string, currentStatus: string) => {
     try {
       const newStatus = (currentStatus === 'inactive') ? 'approved' : 'inactive';
       await updateDoc(doc(db, 'users', id), { status: newStatus });
     } catch (err: any) {
       alert("שגיאה בשינוי סטטוס משתמש: " + err.message);
     }
  };

  const approveUser = async (id: string) => {
     try {
       await updateDoc(doc(db, 'users', id), { status: 'approved' });
     } catch (err: any) {
       alert("שגיאה באישור משתמש: " + err.message);
     }
  };

  return (
    <div className="max-w-3xl mx-auto space-y-6">
      
      <div className="bg-gradient-to-l from-indigo-50 to-white rounded-xl shadow-sm border border-indigo-100 p-6 flex items-center justify-between flex-wrap gap-4">
         <div>
            <h3 className="text-lg font-bold text-slate-800 mb-1 flex items-center gap-2">
              <Building size={20} className="text-indigo-500" /> עריכת פרטי חברה המופיעים בקטלוג
            </h3>
            <p className="text-xs text-slate-500">לוגו, כתובות פרסום ופרטי התקשרות</p>
         </div>
         <button 
           onClick={onNavigateToCompany}
           className="px-4 py-2 bg-white border border-indigo-200 text-indigo-700 rounded-md font-medium text-sm hover:bg-indigo-50 transition-colors shadow-sm"
         >
           עבור להגדרות חברה
         </button>
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 overflow-hidden">
        <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
          <Shield size={20} className="text-rose-500" />
          ניהול משתמשים והרשאות
        </h3>

        <div className="flex flex-col gap-3 mb-6 bg-slate-50 p-4 rounded-lg border border-slate-200">
          <div className="flex flex-col md:flex-row gap-3">
            <input 
              type="email" 
              placeholder="אימייל משתמש" 
              value={newUserEmail} 
              onChange={e => setNewUserEmail(e.target.value)}
              className="flex-1 px-3 py-2 border border-slate-200 rounded-md text-[13px] outline-none"
            />
            <input 
              type="text" 
              placeholder="שם מלא / חברה" 
              value={newUserCompany} 
              onChange={e => setNewUserCompany(e.target.value)}
              className="flex-1 px-3 py-2 border border-slate-200 rounded-md text-[13px] outline-none"
            />
            <input 
              type="text" 
              placeholder="פלאפון" 
              value={newUserPhone} 
              onChange={e => setNewUserPhone(e.target.value)}
              className="flex-1 px-3 py-2 border border-slate-200 rounded-md text-[13px] outline-none"
            />
            <select 
              value={newUserRole}
              onChange={e => setNewUserRole(e.target.value)}
              className="px-3 py-2 border border-slate-200 rounded-md text-[13px] outline-none bg-white font-medium"
            >
              <option value="admin">מנהל (Admin)</option>
              <option value="editor">עורך (Editor)</option>
              <option value="viewer">צופה (Viewer)</option>
            </select>
          </div>
          
          <div className="text-[12px] text-slate-600 font-medium mb-1">תפריטים מורשים:</div>
          <div className="flex flex-wrap gap-2">
            {MENU_OPTIONS.map(m => (
              <label key={m.id} className="flex items-center gap-1.5 bg-white border border-slate-200 px-2 py-1.5 rounded-md text-[12px] cursor-pointer hover:bg-slate-50">
                <input 
                  type="checkbox" 
                  checked={newUserMenus.includes(m.id)}
                  onChange={() => toggleMenu(m.id, newUserMenus, setNewUserMenus)}
                  className="accent-indigo-600"
                />
                {m.label}
              </label>
            ))}
          </div>

          <div className="text-xs text-amber-700 bg-amber-50 p-2 rounded-md mt-2 border border-amber-200">
             * שים לב: הוספת משתמש מגדירה את הרשאותיו מראש. המשתמש יצטרך לבצע <b>הרשמה / צור חשבון במסך ההתחברות</b> עם אותו אימייל בדיוק כדי שאלו יחולו.
          </div>

          <button 
            onClick={handleAddUser}
            className="self-end flex items-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2 mt-2 rounded-md text-[13px] font-bold transition-colors w-full md:w-auto justify-center"
          >
            <UserPlus size={16} /> הוסף משתמש
          </button>
        </div>

        <div className="border border-slate-200 rounded-md overflow-x-auto">
          <table className="w-full text-[13px] text-right min-w-[600px]">
            <thead className="bg-slate-50 border-b border-slate-200 text-slate-500 font-bold">
              <tr>
                <th className="px-4 py-3">כתובת אימייל</th>
                <th className="px-4 py-3">חברה / פלאפון</th>
                <th className="px-4 py-3">הרשאה</th>
                <th className="px-4 py-3">תפריטים</th>
                <th className="px-4 py-3 text-center">פעולות</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 font-medium">
              {users.map(u => (
                <tr key={u.id} className="hover:bg-slate-50">
                  <td className="px-4 py-3 text-slate-800">
                    <div className="flex items-center gap-2">
                       {u.email}
                       {u.status === 'pending' && <span className="px-2 py-0.5 bg-amber-100 text-amber-700 rounded-md text-[10px] font-bold shrink-0">ממתין</span>}
                       {u.status === 'inactive' && <span className="px-2 py-0.5 bg-red-100 text-red-700 rounded-md text-[10px] font-bold shrink-0">לא פעיל</span>}
                    </div>
                  </td>
                  
                  {editingUserId === u.id ? (
                    <td className="px-4 py-3" colSpan={3}>
                      <div className="flex flex-col gap-3">
                        <div className="flex gap-2">
                          <input 
                            type="text" 
                            placeholder="שם חברה"
                            value={editCompany}
                            onChange={e => setEditCompany(e.target.value)}
                            className="px-2 py-1 border border-slate-300 rounded-md outline-none bg-white w-full max-w-[150px]"
                          />
                          <input 
                            type="text" 
                            placeholder="פלאפון"
                            value={editPhone}
                            onChange={e => setEditPhone(e.target.value)}
                            className="px-2 py-1 border border-slate-300 rounded-md outline-none bg-white w-full max-w-[150px]"
                          />
                        </div>
                        <select 
                          value={editRole}
                          onChange={e => setEditRole(e.target.value)}
                          className="px-2 py-1 border border-slate-300 rounded-md outline-none bg-white mb-2 w-full max-w-[200px]"
                        >
                          <option value="admin">מנהל</option>
                          <option value="editor">עורך</option>
                          <option value="viewer">צופה</option>
                        </select>
                        <div className="flex flex-wrap gap-2">
                          {MENU_OPTIONS.map(m => (
                            <label key={m.id} className="flex items-center gap-1.5 bg-white border border-slate-200 px-2 py-1 rounded-md text-[11px] cursor-pointer">
                              <input 
                                type="checkbox" 
                                checked={editMenus.includes(m.id)}
                                onChange={() => toggleMenu(m.id, editMenus, setEditMenus)}
                                className="accent-indigo-600"
                              />
                              {m.label}
                            </label>
                          ))}
                        </div>
                      </div>
                    </td>
                  ) : (
                    <>
                      <td className="px-4 py-3 text-slate-600 text-xs text-right">
                        <div>{u.company || '—'}</div>
                        <div className="text-slate-400 mt-0.5">{u.phone || '—'}</div>
                      </td>
                      <td className="px-4 py-3">
                        <span className={`px-2 py-1 rounded-full text-[11px] ${
                          u.role === 'admin' ? 'bg-rose-100 text-rose-700' :
                          u.role === 'editor' ? 'bg-indigo-100 text-indigo-700' :
                          'bg-slate-200 text-slate-700'
                        }`}>
                          {u.role === 'admin' ? 'מנהל' : u.role === 'editor' ? 'עורך' : 'צופה'}
                        </span>
                      </td>
                      <td className="px-4 py-3">
                        <div className="flex flex-wrap gap-1">
                          {(u.menus || []).map((mId: string) => {
                            const found = MENU_OPTIONS.find(o => o.id === mId);
                            return found ? <span key={mId} className="bg-slate-100 border border-slate-200 text-slate-600 px-1.5 py-0.5 rounded text-[10px]">{found.label}</span> : null;
                          })}
                          {(!u.menus || u.menus.length === 0) && <span className="text-slate-400 text-xs">ברירת מחדל</span>}
                        </div>
                      </td>
                    </>
                  )}

                  <td className="px-4 py-3 text-center">
                    {editingUserId === u.id ? (
                      <div className="flex items-center justify-center gap-1">
                        <button onClick={() => saveEditUser(u.id)} className="text-emerald-600 hover:bg-emerald-50 p-1.5 rounded-md transition-colors" title="שמור">
                          <Check size={16} />
                        </button>
                        <button onClick={() => setEditingUserId(null)} className="text-slate-500 hover:bg-slate-100 p-1.5 rounded-md transition-colors" title="ביטול">
                          <X size={16} />
                        </button>
                      </div>
                    ) : (
                      <div className="flex items-center justify-center gap-1 relative">
                        {u.status === 'pending' ? (
                          <button onClick={() => approveUser(u.id)} className="text-emerald-600 bg-emerald-50 hover:bg-emerald-100 px-2 py-1.5 rounded-md transition-colors text-[11px] font-bold flex items-center gap-1" title="אשר משתמש">
                            <Check size={14} /> אשר מסמך
                          </button>
                        ) : (
                          <button 
                            onClick={() => u.role !== 'admin' && changeUserStatus(u.id, u.status || 'approved')} 
                            disabled={u.role === 'admin'}
                            className={`min-w-[70px] ${u.status === 'inactive' ? 'text-rose-600 bg-rose-50 hover:bg-rose-100 border-rose-200' : 'text-emerald-600 bg-emerald-50 hover:bg-emerald-100 border-emerald-200'} border px-2 py-1.5 rounded-md transition-colors text-[11px] font-bold flex items-center justify-center gap-1 ${u.role === 'admin' ? 'opacity-50 cursor-not-allowed' : ''}`}
                            title={u.role === 'admin' ? "לא ניתן להשעות מנהל" : ""}
                          >
                            {u.status === 'inactive' ? 'לא פעיל' : 'פעיל'}
                          </button>
                        )}
                        <button onClick={() => handleResetPassword(u.email)} className="text-indigo-500 hover:bg-indigo-50 p-1.5 rounded-md transition-colors" title="שלח קישור לאיפוס סיסמא">
                          <Key size={16} />
                        </button>
                        <button onClick={() => startEditUser(u)} className="text-slate-500 hover:bg-slate-100 p-1.5 rounded-md transition-colors" title="ערוך משתמש">
                          <Edit size={16} />
                        </button>
                        {userToDelete === u.id ? (
                           <>
                             <div className="absolute bg-white border border-slate-200 shadow-lg rounded-lg p-2 flex flex-col gap-2 z-10 bottom-full mb-1 ml-4 whitespace-nowrap">
                               <div className="text-[11px] font-bold text-slate-700">בטוח למחוק?</div>
                               <div className="flex gap-1">
                                  <button onClick={() => handleDelUser(u.id)} className="bg-red-600 text-white px-2 py-1 text-[10px] rounded hover:bg-red-700">מחק</button>
                                  <button onClick={() => setUserToDelete(null)} className="bg-slate-200 text-slate-800 px-2 py-1 text-[10px] rounded hover:bg-slate-300">בטל</button>
                               </div>
                             </div>
                           </>
                        ) : null}
                        <button onClick={() => setUserToDelete(u.id)} className="text-red-500 hover:bg-red-50 p-1.5 rounded-md transition-colors" title="מחק משתמש">
                          <Trash2 size={16} />
                        </button>
                      </div>
                    )}
                  </td>
                </tr>
              ))}
              {users.length === 0 && (
                <tr>
                  <td colSpan={4} className="px-4 py-8 text-center text-slate-500">
                    עדיין לא נוספו הרשאות מיוחדות.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
