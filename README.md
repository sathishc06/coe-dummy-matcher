# COE Dummy Matcher — Mobile build package

Includes a mobile-responsive Streamlit web app and an Android Studio WebView wrapper project. The Android wrapper loads the securely hosted web app; examination ZIPs are processed on that server, not on the phone.

## Web app
1. Deploy `web/` to an institution-approved HTTPS Streamlit host.
2. Install dependencies from `requirements.txt` and run `streamlit run app.py`.
3. On Android Chrome, open the HTTPS URL → ⋮ → **Add to Home screen** / **Install app**.

## Android APK
1. Open `android/` in Android Studio (JDK 17, Android SDK 35).
2. In `android/app/src/main/java/in/coe/matcher/MainActivity.java`, replace `https://YOUR-COE-APP-HOST.example` with your deployed HTTPS URL.
3. Sync Gradle, then **Build → Build APK(s)**. Output will be under `android/app/build/outputs/apk/debug/`.

This environment has no Android SDK/Gradle installation, so it cannot compile or certify an APK here. The complete Android wrapper source is provided. Do not deploy confidential exam data to a public or unapproved host. The current matching engine still requires independent date/session/index validation before official release; this mobile package is a UI/deployment scaffold, not a certification of the matching logic.
