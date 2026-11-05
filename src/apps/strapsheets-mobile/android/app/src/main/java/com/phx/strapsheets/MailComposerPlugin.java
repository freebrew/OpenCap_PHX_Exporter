package com.phx.strapsheets;

import android.content.ClipData;
import android.content.Intent;
import android.net.Uri;
import androidx.core.content.FileProvider;
import com.getcapacitor.JSArray;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;
import java.io.File;
import java.util.ArrayList;
import org.json.JSONException;

/** Opens the phone's mail app with recipients, subject, and PDF attachments filled in. */
@CapacitorPlugin(name = "MailComposer")
public class MailComposerPlugin extends Plugin {
    @PluginMethod
    public void compose(PluginCall call) {
        JSArray recipients = call.getArray("to");
        JSArray files = call.getArray("files");
        String subject = call.getString("subject", "");
        String text = call.getString("text", "");
        if (files == null || files.length() == 0) {
            call.reject("Choose a strap sheet to email.");
            return;
        }
        if (getActivity() == null) {
            call.reject("The mail app could not be opened.");
            return;
        }
        try {
            ArrayList<Uri> uris = new ArrayList<>();
            for (int index = 0; index < files.length(); index++) {
                uris.add(shareable(files.getString(index)));
            }
            Intent intent = new Intent(Intent.ACTION_SEND_MULTIPLE);
            intent.setType("application/pdf");
            intent.putExtra(Intent.EXTRA_EMAIL, strings(recipients));
            intent.putExtra(Intent.EXTRA_SUBJECT, subject);
            intent.putExtra(Intent.EXTRA_TEXT, text);
            intent.putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris);
            intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
            intent.setClipData(clip(uris));
            Intent chooser = Intent.createChooser(intent, "Send strap sheet");
            chooser.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
            chooser.setClipData(intent.getClipData());
            getActivity().runOnUiThread(() -> {
                try {
                    getActivity().startActivity(chooser);
                    call.resolve();
                } catch (RuntimeException error) {
                    call.reject("The mail app could not be opened.", error);
                }
            });
        } catch (Exception error) {
            call.reject("The mail app could not be opened.", error);
        }
    }

    private Uri shareable(String raw) {
        Uri uri = Uri.parse(raw);
        if (uri.getScheme() == null || !"file".equals(uri.getScheme())) return uri;
        String path = uri.getPath();
        if (path == null) return uri;
        return FileProvider.getUriForFile(getContext(), getContext().getPackageName() + ".fileprovider", new File(path));
    }

    private String[] strings(JSArray values) throws JSONException {
        if (values == null) return new String[0];
        String[] out = new String[values.length()];
        for (int index = 0; index < out.length; index++) out[index] = values.getString(index);
        return out;
    }

    private ClipData clip(ArrayList<Uri> uris) {
        ClipData clip = ClipData.newUri(getContext().getContentResolver(), "Strap sheet", uris.get(0));
        for (int index = 1; index < uris.size(); index++) clip.addItem(new ClipData.Item(uris.get(index)));
        return clip;
    }
}
