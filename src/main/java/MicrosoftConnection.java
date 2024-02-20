import com.google.gson.Gson;
import com.google.gson.JsonElement;
import com.google.gson.reflect.TypeToken;
import com.microsoft.aad.msal4j.*;
import com.microsoft.aad.msal4j.Prompt;
import com.microsoft.aad.msal4j.PublicClientApplication;
import com.microsoft.graph.models.*;
import com.microsoft.graph.options.FunctionOption;
import com.microsoft.graph.requests.*;
import okhttp3.Request;

import java.io.InputStream;
import java.lang.reflect.Type;
import java.net.URI;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.*;
import java.util.List;
import java.util.concurrent.CompletableFuture;

public class MicrosoftConnection {
    private final Set<String> scope = new HashSet<>(Arrays.asList("user.read", "files.read.all", "offline_access", "sites.readwrite.all",
            "ChatMessage.Read", "Files.ReadWrite.All", "Chat.ReadWrite"));
    private final String clientId;
    private final String clientSecret;
    private final String tenant;
    private final Set<String> scope2 = new HashSet<>(Arrays.asList(".default"));
    private final String code;
    private String tokenCache;
    private IAuthenticationResult iAuthenticationResult;

    public MicrosoftConnection(Properties properties){
        this.clientId = properties.getProperty("microsoft.client.id");
        this.clientSecret = properties.getProperty("microsoft.client.secret");
        this.tenant = properties.getProperty("microsoft.tenant");
        this.code = properties.getProperty("microsoft.authorization.code");
    }

    public void connection1() throws Exception {
        PublicClientApplication clientApplication = PublicClientApplication.builder(clientId).build();
        InteractiveRequestParameters parameters = InteractiveRequestParameters.builder(new URI("http://localhost:8777"))
                .scopes(scope)
                .prompt(Prompt.SELECT_ACCOUNT)
                .httpPollingTimeoutInSeconds(20)
                .build();

        iAuthenticationResult = clientApplication.acquireToken(parameters).get();
        tokenCache = clientApplication.tokenCache().serialize();
        // tokenCache и iAuthenticationResult необходимо сохранить в БД
    }

    public void connection2() throws Exception {
        String url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?" +
                "client_id=" + clientId +
                "&redirect_uri=http://localhost:8777" +
                "&scope=offline_access+user.read+files.read.all+files.readwrite" +
                "&response_type=code" +
                "&response_mode=fragment" +
                "&prompt=select_account";
        //Получение кода авторизации

        PublicClientApplication clientApplication = PublicClientApplication.builder(clientId).build();
        AuthorizationCodeParameters parameters = AuthorizationCodeParameters
                .builder(code, new URI("http://localhost:8777"))
                .scopes(scope)
                .build();

        iAuthenticationResult = clientApplication.acquireToken(parameters).get();
        tokenCache = clientApplication.tokenCache().serialize();
        // tokenCache и iAuthenticationResult необходимо сохранить в БД
    }

    public void connection3() throws Exception {
        ConfidentialClientApplication clientApplication = ConfidentialClientApplication.builder(clientId,
                ClientCredentialFactory.createFromSecret(clientSecret))
                .build();
        ClientCredentialParameters parameters = ClientCredentialParameters.builder(scope2)
                .tenant(tenant)
                .build();
        iAuthenticationResult = clientApplication.acquireToken(parameters).get();
        tokenCache = clientApplication.tokenCache().serialize();
    }

    public void connection4() throws Exception {
        PublicClientApplication clientApplication = PublicClientApplication.builder(clientId).build();
        UserNamePasswordParameters parameters = UserNamePasswordParameters
                .builder(scope, "username@mail.com", "password".toCharArray())
                .tenant(tenant)
                .build();
        iAuthenticationResult = clientApplication.acquireToken(parameters).get();
        tokenCache = clientApplication.tokenCache().serialize();
    }

    public void reloadToken() throws Exception {
        PublicClientApplication clientApplication = PublicClientApplication.builder(clientId).build();
        clientApplication.tokenCache().deserialize(tokenCache);
        IAccount iAccount = iAuthenticationResult.account();
        SilentParameters parameters = SilentParameters.builder(scope, iAccount)
                .build();

        long date = iAuthenticationResult.expiresOnDate().getTime();
        long now = new Date().getTime();
        if (now > date) {
            iAuthenticationResult = clientApplication.acquireTokenSilently(parameters).get();
        }

        tokenCache = clientApplication.tokenCache().serialize();
        // tokenCache и iAuthenticationResult необходимо сохранить в БД
    }

    public List<List<String>> getListTableRows() {
        String token = iAuthenticationResult.accessToken();
        GraphServiceClient<Request> graphClient = GraphServiceClient.builder()
                .authenticationProvider(requestUrl -> CompletableFuture.completedFuture(token))
                .buildClient();
        DriveItemCollectionPage driveItemCollectionPage = graphClient.me().drive().root().children().buildRequest().get();
        List<List<String>> resultTableRow = new ArrayList();
        for (DriveItem driveItem : driveItemCollectionPage.getCurrentPage()) {
            if (null != driveItem.file && driveItem.name.equals("Book.xlsx")) {
                List<WorkbookWorksheet> currentPage = graphClient.me().drive().items(driveItem.id).workbook().worksheets().buildRequest().get().getCurrentPage();
                List<WorkbookTable> workbookTable = graphClient.me().drive().items(driveItem.id).workbook().worksheets(currentPage.get(0).id)
                        .tables().buildRequest().get().getCurrentPage();
                List<WorkbookTableRow> tableRow = graphClient.me().drive().items(driveItem.id).workbook().worksheets(currentPage.get(0).id)
                        .tables(workbookTable.get(0).id).rows().buildRequest().get().getCurrentPage();
                for (WorkbookTableRow workbookTableRow : tableRow) {
                    JsonElement values = workbookTableRow.values;
                    Type collectionType = new TypeToken<Collection<List<String>>>() {}.getType();
                    Collection<List<String>> list = new Gson().fromJson(values, collectionType);
                    resultTableRow.addAll(list);
                }
            }
        }
        return resultTableRow;
    }

    public List<List<String>> getListWorkSheetRows() {
        String token = iAuthenticationResult.accessToken();
        GraphServiceClient<Request> graphClient = GraphServiceClient.builder()
                .authenticationProvider(requestUrl -> CompletableFuture.completedFuture(token))
                .buildClient();

        List<List<String>> resultWorkSheetRow = new ArrayList();
        DriveItemCollectionPage driveItemCollectionPage = graphClient.me().drive().root().children().buildRequest().get();
        for (DriveItem driveItem : driveItemCollectionPage.getCurrentPage()) {
            if (null != driveItem.file && driveItem.name.equals("Book.xlsx")) {
                List<WorkbookWorksheet> currentPage = graphClient.me().drive().items(driveItem.id).workbook().worksheets().buildRequest().get().getCurrentPage();
                WorkbookRange workbookRange = graphClient.me().drive().items(driveItem.id).workbook().worksheets(currentPage.get(0).id)
                        .range().buildRequest(new FunctionOption("address", "A1:B10")).get();

                JsonElement values = workbookRange.values;
                Type collectionType = new TypeToken<Collection<List<String>>>() {}.getType();
                Collection<List<String>> list = new Gson().fromJson(values, collectionType);
                resultWorkSheetRow.addAll(list);
            }
        }
        return resultWorkSheetRow;
    }

    public void getVideo() {
        String token = iAuthenticationResult.accessToken();
        GraphServiceClient<Request> graphClient = GraphServiceClient.builder()
                .authenticationProvider(requestUrl -> CompletableFuture.completedFuture(token))
                .buildClient();

        String meetingId = "19:meeting_NGY2NzFiMmMtOGU4My00NGI5LTkwYzItNDJlMzRlZmNkYWZl@thread.v2";
        ChatMessageCollectionPage page = graphClient.me().chats(meetingId).messages().buildRequest().get();
//        ChatMessageCollectionPage page = graphClient.chats(meetingId).messages()
//                .buildRequest(new Option[]{new QueryOption("top", 10)}).get();
        ChatMessageCollectionRequestBuilder nextPage = page.getNextPage();
        for (ChatMessage chatMessage : page.getCurrentPage()) {
            if (chatMessage.eventDetail instanceof CallRecordingEventMessageDetail) {
                CallRecordingEventMessageDetail callRecording = (CallRecordingEventMessageDetail) chatMessage.eventDetail;
                if (callRecording.callRecordingStatus == CallRecordingStatus.SUCCESS) {
                    String callRecordingUrl = callRecording.callRecordingUrl;
                    String base64 = Base64.getUrlEncoder().encodeToString(callRecordingUrl.getBytes(StandardCharsets.UTF_8));
                    String encode = "u!" + base64.replaceAll("/", "_").replaceAll("/+", "-");
                    try {
                        InputStream inputStream1 = graphClient.shares(encode).driveItem().content().buildRequest().get();
                        Files.copy(inputStream1, Paths.get("C:\\" + callRecording.callRecordingDisplayName), StandardCopyOption.REPLACE_EXISTING);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            }
        }
//        OnlineMeetingCollectionPage onlineMeetingCollectionPage = graphClient.me().onlineMeetings().buildRequest().filter(String.format("joinWebUrl eq '%s'", "https://teams.microsoft.com/l/meetup-join/19%3ameeting_NGY2NzFiMmMtOGU4My00NGI5LTkwYz.v2/0")).get();
    }
}

