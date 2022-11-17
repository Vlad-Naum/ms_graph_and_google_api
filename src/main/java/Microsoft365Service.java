import com.google.common.cache.Cache;
import com.google.common.cache.CacheBuilder;
import com.microsoft.aad.msal4j.*;
import com.microsoft.graph.models.User;
import com.microsoft.graph.requests.DriveItemCollectionPage;
import com.microsoft.graph.requests.DriveSharedWithMeCollectionPage;
import com.microsoft.graph.requests.GraphServiceClient;
import okhttp3.Request;

import java.net.URI;
import java.net.URISyntaxException;
import java.time.Duration;
import java.util.*;
import java.util.concurrent.CompletableFuture;

public class Microsoft365Service {
    private static final Cache<String, String> cacheToken = CacheBuilder.newBuilder().expireAfterWrite(Duration.ofHours(1)).build();
    private final Set<String> scope = new HashSet<>(Arrays.asList("user.read", "files.read.all", "offline_access"));
    private final String clientId;

    public Microsoft365Service(Properties properties) {
        this.clientId = properties.getProperty("microsoft.client.id");
    }

    public PublicClientApplication getMicrosoftApplication(String tokenCache) {
        PublicClientApplication clientApplication = PublicClientApplication.builder(clientId).build();
        if (tokenCache != null) {
            clientApplication.tokenCache().deserialize(tokenCache);
        }
        return clientApplication;
    }

    public AuthorizationCodeParameters getAuthorizationCodeParameters(String code) throws URISyntaxException {
        return AuthorizationCodeParameters
                .builder(code, new URI("http://localhost:8777"))
                .scopes(scope)
                .build();
    }

    public GraphServiceClient<Request> getMicrosoftGraphClient(IAuthenticationResult iAuthenticationResult, String tokenCache) throws Exception {
        long date = iAuthenticationResult.expiresOnDate().getTime();
        long now = new Date().getTime();
        IAccount iAccount = iAuthenticationResult.account();
        if (now > date) {
            SilentParameters parameters = SilentParameters.builder(scope, iAccount)
                    .build();
            iAuthenticationResult = getMicrosoftApplication(tokenCache).acquireTokenSilently(parameters).get();
        }
        final String token = iAuthenticationResult.accessToken();
        String cacheToken = Microsoft365Service.cacheToken.getIfPresent(iAccount.username());
        if (!token.equals(cacheToken)) {
            Microsoft365Service.cacheToken.put(iAccount.username(), token);
        }
        return GraphServiceClient.builder()
                .authenticationProvider(requestUrl -> CompletableFuture.completedFuture(token))
                .buildClient();
    }

    public User getUser(GraphServiceClient<Request> graphClient) {
        User user = graphClient.me().buildRequest().get();
        if (user == null) {
            // Прокидываем исключение
        }
        return user;
    }

    public DriveItemCollectionPage getDriveItemCollectionPage(GraphServiceClient<Request> graphClient, String driveItemId) {
        DriveItemCollectionPage driveItemCollectionPage;
        if (driveItemId == null) {
            driveItemCollectionPage = graphClient.me().drive().root().children().buildRequest().get();
        } else {
            driveItemCollectionPage = graphClient.me().drive().items(driveItemId).children().buildRequest().get();
        }
        return driveItemCollectionPage;
    }

    public DriveSharedWithMeCollectionPage getDriveSharedCollectionPage(GraphServiceClient<Request> graphClient) {
        return graphClient.me().drive().sharedWithMe().buildRequest().get();
    }
}
