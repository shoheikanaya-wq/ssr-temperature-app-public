package jp.shoheikanaya.ssrtemperature;

import android.annotation.SuppressLint;
import android.content.ContentValues;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.provider.MediaStore;
import android.util.Base64;
import android.webkit.JavascriptInterface;
import android.webkit.SafeBrowsingResponse;
import android.webkit.WebResourceRequest;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;
import androidx.webkit.WebViewCompat;
import androidx.webkit.WebViewFeature;

import java.io.OutputStream;

public class MainActivity extends AppCompatActivity {
  private static final String APP_URL = "https://shoheikanaya-wq.github.io/ssr-temperature-app-public/";
  private static final String APP_HOST = "shoheikanaya-wq.github.io";
  private WebView webView;

  @SuppressLint({"SetJavaScriptEnabled", "JavascriptInterface"})
  @Override protected void onCreate(Bundle savedInstanceState) {
    super.onCreate(savedInstanceState);

    webView = new WebView(this);
    setContentView(webView);

    webView.getSettings().setJavaScriptEnabled(true);
    webView.getSettings().setDomStorageEnabled(true);
    webView.getSettings().setAllowFileAccess(false);
    webView.getSettings().setAllowContentAccess(false);
    webView.getSettings().setMixedContentMode(android.webkit.WebSettings.MIXED_CONTENT_NEVER_ALLOW);
    webView.getSettings().setSavePassword(false);

    android.webkit.CookieManager.getInstance().setAcceptThirdPartyCookies(webView, false);
    webView.addJavascriptInterface(new DownloadBridge(), "SSRAndroid");

    webView.setWebViewClient(new WebViewClient() {
      @Override public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest request) {
        Uri u = request.getUrl();
        if ("https".equalsIgnoreCase(u.getScheme()) && APP_HOST.equalsIgnoreCase(u.getHost())) return false;
        startActivity(new Intent(Intent.ACTION_VIEW, u));
        return true;
      }

      @Override public void onPageFinished(WebView view, String url) {
        super.onPageFinished(view, url);
        injectDownloadBridge();
      }
    });

    if (WebViewFeature.isFeatureSupported(WebViewFeature.SAFE_BROWSING_ENABLE)) {
      WebViewCompat.startSafeBrowsing(this, value -> {});
    }

    webView.loadUrl(APP_URL);
  }

  private void injectDownloadBridge() {
    String js =
      "(function(){if(window.__ssrAndroidReady)return;window.__ssrAndroidReady=true;" +
      "document.addEventListener('click',function(e){var a=e.target.closest&&e.target.closest('a[download]');" +
      "if(!a||!a.href||a.href.indexOf('blob:')!==0)return;e.preventDefault();" +
      "var name=a.getAttribute('download')||'SSR温測_生成.xlsx';" +
      "fetch(a.href).then(function(r){return r.blob();}).then(function(b){var fr=new FileReader();" +
      "fr.onloadend=function(){SSRAndroid.saveBase64(name,String(fr.result).split(',')[1]||'');};" +
      "fr.readAsDataURL(b);});},true);})();";
    webView.evaluateJavascript(js, null);
  }

  public class DownloadBridge {
    @JavascriptInterface public void saveBase64(String fileName, String base64) {
      runOnUiThread(() -> {
        try {
          byte[] bytes = Base64.decode(base64, Base64.DEFAULT);
          String safeName = (fileName == null || fileName.trim().isEmpty()) ? "SSR温測_生成.xlsx" : fileName.replaceAll("[\\\\/:*?\"<>|]", "_");
          ContentValues cv = new ContentValues();
          cv.put(MediaStore.Downloads.DISPLAY_NAME, safeName);
          cv.put(MediaStore.Downloads.MIME_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
          cv.put(MediaStore.Downloads.IS_PENDING, 1);
          Uri uri = getContentResolver().insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, cv);
          if (uri == null) throw new IllegalStateException("保存先を作成できません");
          try (OutputStream os = getContentResolver().openOutputStream(uri)) {
            if (os == null) throw new IllegalStateException("保存先を開けません");
            os.write(bytes);
          }
          cv.clear();
          cv.put(MediaStore.Downloads.IS_PENDING, 0);
          getContentResolver().update(uri, cv, null, null);
          Toast.makeText(MainActivity.this, "Excelをダウンロードに保存しました", Toast.LENGTH_LONG).show();
        } catch (Exception ex) {
          Toast.makeText(MainActivity.this, "Excel保存エラー: " + ex.getMessage(), Toast.LENGTH_LONG).show();
        }
      });
    }
  }

  @Override public void onBackPressed() {
    if (webView != null && webView.canGoBack()) webView.goBack();
    else super.onBackPressed();
  }
}
