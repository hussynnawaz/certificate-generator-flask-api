import firebase_admin
from firebase_admin import credentials, firestore, auth

# Initialize Firebase Admin SDK with your Firebase project credentials
cred = credentials.Certificate('firebaseConfig.json')
firebase_admin.initialize_app(cred)

# Access Firestore and Authentication
db = firestore.client()
auth = firebase_admin.auth

# Example to get a user from Firebase Auth
def get_user(uid):
    try:
        user = auth.get_user(uid)
        print(f'User found: {user.uid}')
    except firebase_admin.auth.UserNotFoundError:
        print('User not found.')

# Example to interact with Firestore (get a document)
def get_document(collection, doc_id):
    doc_ref = db.collection(collection).document(doc_id)
    doc = doc_ref.get()
    if doc.exists:
        print(f'Document data: {doc.to_dict()}')
    else:
        print('No such document!')

# Example of calling the functions
get_user('some-user-uid')
get_document('users', 'some-doc-id')
