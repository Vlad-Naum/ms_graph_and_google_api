import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.auth.oauth2.StoredCredential;
import com.google.api.client.auth.oauth2.TokenResponse;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.DataStore;
import com.google.api.client.util.store.MemoryDataStoreFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.File;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;

import java.io.*;
import java.security.GeneralSecurityException;
import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.List;

public class GoogleConnection {
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS_READONLY);
    private static final String clientId = "";
    private static final String clientSecret = "";
    private static final String redirectURI = "http://localhost";
    private DataStore<StoredCredential> credentialDataStore;

    private Credential firstConnection() throws IOException, GeneralSecurityException {
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        GoogleClientSecrets.Details details = new GoogleClientSecrets.Details()
                .setClientId(clientId)
                .setClientSecret(clientSecret)
                .setRedirectUris(Collections.singletonList(redirectURI));
        GoogleClientSecrets clientSecrets = new GoogleClientSecrets().setInstalled(details);
        MemoryDataStoreFactory memoryDataStoreFactory = new MemoryDataStoreFactory();

        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(memoryDataStoreFactory)
                .setAccessType("offline").build();

        TokenResponse tokenResponse = flow
                .newTokenRequest("4/0AfgeXvuJ8NFmYnLnQbOcIll07NlE4WvAPrZeMYSfIhobYMBWvEwZp2Us2nVhove4dKfHFg")
                .setRedirectUri(redirectURI).execute();
        Credential credential = flow.createAndStoreCredential(tokenResponse, "user_id");
        credentialDataStore = flow.getCredentialDataStore();
        return credential;
    }

    private Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        GoogleClientSecrets.Details details = new GoogleClientSecrets.Details()
                .setClientId(clientId)
                .setClientSecret(clientSecret)
                .setRedirectUris(Collections.singletonList(redirectURI));
        GoogleClientSecrets clientSecrets = new GoogleClientSecrets().setInstalled(details);
        MemoryDataStoreFactory memoryDataStoreFactory = new MemoryDataStoreFactory();

        GoogleAuthorizationCodeFlow.Builder builder = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(memoryDataStoreFactory)
                .setAccessType("offline");
        if (credentialDataStore != null) {
            builder.setCredentialDataStore(credentialDataStore);
        }
        GoogleAuthorizationCodeFlow flow = builder.build();

        Credential credential = flow.loadCredential("user_id");
        if (credential == null) {
            return null;
        }
        long now = new Date().getTime();
        if (credential.getExpiresInSeconds() != null && credential.getExpirationTimeMilliseconds() < now) {
            credential.refreshToken();
            credentialDataStore = flow.getCredentialDataStore();
            //Сохранять в БД credentialDataStore
        }

        return credential;
    }

    public void connection() throws IOException, GeneralSecurityException {
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        Credential credential = firstConnection();
        Drive drive = new Drive.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential).build();
        List<File> list = drive.files().list().setQ("sharedWithMe=true").execute().getFiles();

        Sheets sheets = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential).build();
        for (File file : list) {
            if (file.getMimeType().contains("spreadsheet")) {
                List<List<Object>> values = sheets.spreadsheets().values()
                        .get(file.getId(), "A1:B10")
                        .execute().getValues();
                System.out.println(values.get(0).get(0).toString());
                values.stream().flatMap(Collection::stream).forEach(System.out::println);
            }
        }
    }
}

