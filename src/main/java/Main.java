import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.reflect.TypeToken;
import com.microsoft.aad.msal4j.AuthorizationCodeParameters;
import com.microsoft.aad.msal4j.IAuthenticationResult;
import com.microsoft.aad.msal4j.PublicClientApplication;
import com.microsoft.graph.models.*;
import com.microsoft.graph.options.FunctionOption;
import com.microsoft.graph.requests.*;
import okhttp3.Request;

import java.io.FileInputStream;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Properties;
import java.util.function.Predicate;
import java.util.stream.Collectors;

public class Main {
    private IAuthenticationResult iAuthenticationResult;
    private String tokenCache;

    public static void main(String[] args) throws Exception {
        FileInputStream fis = new FileInputStream("src/main/resources/resources.properties");
        Properties property = new Properties();
        property.load(fis);

        MicrosoftConnection connection = new MicrosoftConnection(property);
        connection.connection1();
        List<List<String>> listTableRows = connection.getListTableRows();

//        GoogleConnection connection = new GoogleConnection(property);
//        connection.firstConnection(property);
    }

    public void createConnection(Properties property) throws Exception {
        // Prepare
        Microsoft365Service service = new Microsoft365Service(property);

        //Execute
        PublicClientApplication clientApplication = service.getMicrosoftApplication(null);
        AuthorizationCodeParameters authorizationCodeParameters = service.getAuthorizationCodeParameters("");
        IAuthenticationResult iAuthenticationResult = clientApplication.acquireToken(authorizationCodeParameters).get();
        String tokenCache = clientApplication.tokenCache().serialize();
        this.iAuthenticationResult = iAuthenticationResult;
        this.tokenCache = tokenCache;
        GraphServiceClient<Request> graphClient = service.getMicrosoftGraphClient(iAuthenticationResult, tokenCache);
        User user = service.getUser(graphClient);
        String upn = user.userPrincipalName;
        // tokenCache и iAuthenticationResult необходимо сохранить в БД, еще upn для отображения в профиле подключения
    }

    // Получение расположения (пока только oneDrive)
    public List<Drive> getDrives(Properties property) throws Exception {
        Microsoft365Service service = new Microsoft365Service(property);
        GraphServiceClient<Request> graphClient = service.getMicrosoftGraphClient(iAuthenticationResult, tokenCache);

        DriveCollectionPage driveCollectionPage = graphClient.me().drives().buildRequest().get();
        if (driveCollectionPage != null) {
            return driveCollectionPage.getCurrentPage();
        }
        return new ArrayList<>();
    }

    // Получение файлов с oneDrive
    public List<DriveItem> getMicrosoftDriveItem(String driveItemId, Properties property) throws Exception {
        Microsoft365Service service = new Microsoft365Service(property);
        GraphServiceClient<Request> graphClient = service.getMicrosoftGraphClient(iAuthenticationResult, tokenCache);

        DriveItemCollectionPage driveItemCollectionPage = service.getDriveItemCollectionPage(graphClient, driveItemId);
        if (driveItemCollectionPage != null) {
            // Возвращаем только папки и .xlsx файлы
            Predicate<DriveItem> predicate = driveItem -> driveItem.folder != null ||
                    (driveItem.name != null && driveItem.name.contains(".xlsx"));
            return driveItemCollectionPage.getCurrentPage().stream().filter(predicate).collect(Collectors.toList());
        }
        return new ArrayList<>();
    }

    // Получение файлов с "Общий доступ"
    public List<DriveItem> getMicrosoftSharedDriveItem(Properties property) throws Exception {
        Microsoft365Service service = new Microsoft365Service(property);
        GraphServiceClient<Request> graphClient = service.getMicrosoftGraphClient(iAuthenticationResult, tokenCache);

        DriveSharedWithMeCollectionPage driveSharedItemCollectionPage = service.getDriveSharedCollectionPage(graphClient);
        if (driveSharedItemCollectionPage != null) {
            // Возвращаем только папки и .xlsx файлы
            Predicate<DriveItem> predicate = driveItem -> driveItem.folder != null ||
                    (driveItem.name != null && driveItem.name.contains(".xlsx"));
            return driveSharedItemCollectionPage.getCurrentPage().stream().filter(predicate).collect(Collectors.toList());
        }
        return new ArrayList<>();
    }

    // Получение листов из .xlsx документа
    public List<WorkbookWorksheet> getWorkbookWorksheets(String driveItemId, Properties property) throws Exception {
        Microsoft365Service service = new Microsoft365Service(property);
        GraphServiceClient<Request> graphClient = service.getMicrosoftGraphClient(iAuthenticationResult, tokenCache);

        WorkbookWorksheetCollectionPage worksheetCollectionPage = graphClient.me().drive().items(driveItemId).workbook().worksheets().buildRequest().get();
        if (worksheetCollectionPage != null) {
            return worksheetCollectionPage.getCurrentPage();
        }
        return new ArrayList<>();
    }

    // Получение таблиц
    public List<WorkbookTable> getWorkbookTable(String driveItemId, String workSheetId, Properties property) throws Exception {
        Microsoft365Service service = new Microsoft365Service(property);
        GraphServiceClient<Request> graphClient = service.getMicrosoftGraphClient(iAuthenticationResult, tokenCache);

        WorkbookTableCollectionPage tableCollectionPage = graphClient.me().drive().items(driveItemId).workbook().worksheets(workSheetId)
                .tables().buildRequest().get();
        if (tableCollectionPage != null) {
            return tableCollectionPage.getCurrentPage();
        }
        return new ArrayList<>();
    }

    // Получение колонок
    public List<WorkbookTableColumn> getTableColumn(String driveItemId, String workSheetId, String tableId, Properties property) throws Exception {
        Microsoft365Service service = new Microsoft365Service(property);
        GraphServiceClient<Request> graphClient = service.getMicrosoftGraphClient(iAuthenticationResult, tokenCache);

        WorkbookTableColumnCollectionPage columnCollectionPage = graphClient.me().drive().items(driveItemId).workbook().worksheets(workSheetId)
                .tables(tableId).columns().buildRequest().get();
        if (columnCollectionPage != null) {
            return columnCollectionPage.getCurrentPage();
        }
        return new ArrayList<>();
    }

    // Получение строк с листа
    public List<List<String>> getListWorkSheetRows(String driveItemId, String workSheetId, String range, Properties property) throws Exception {
        Microsoft365Service service = new Microsoft365Service(property);
        GraphServiceClient<Request> graphClient = service.getMicrosoftGraphClient(iAuthenticationResult, tokenCache);
        List<List<String>> resultWorkSheetRow = new ArrayList();

        WorkbookRange workbookRange;
        if (range != null) {
            workbookRange = graphClient.me().drive().items(driveItemId).workbook().worksheets(workSheetId)
                    .range().buildRequest(new FunctionOption("address", range)).get();
        } else {
            workbookRange = graphClient.me().drive().items(driveItemId).workbook().worksheets(workSheetId)
                    .usedRange().buildRequest().get();
        }
        if (workbookRange != null) {
            JsonElement values = workbookRange.values;
            Type collectionType = new TypeToken<Collection<List<String>>>() {}.getType();
            Collection<List<String>> list = new Gson().fromJson(values, collectionType);
            if (list != null) {
                resultWorkSheetRow.addAll(list);
            }
        }
        return resultWorkSheetRow;
    }

    // Получение строк из таблицы
    public List<List<String>> getTableWorkSheetRows(String driveItemId, String workSheetId, String tableId, String columnId, Properties property) throws Exception {
        Microsoft365Service service = new Microsoft365Service(property);
        GraphServiceClient<Request> graphClient = service.getMicrosoftGraphClient(iAuthenticationResult, tokenCache);
        List<List<String>> resultTabletRow = new ArrayList();

        if (columnId != null) {
            WorkbookTableColumn workbookTableColumn = graphClient.me().drive().items(driveItemId).workbook().worksheets(workSheetId)
                    .tables(tableId).columns(columnId).buildRequest().get();
            if (workbookTableColumn != null && workbookTableColumn.values != null) {
                JsonArray jsonArray = workbookTableColumn.values.getAsJsonArray();
                // Под нулевым индексом хранится заголовок колонки, он нам не нужен
                jsonArray.remove(0);
                Type collectionType = new TypeToken<Collection<List<String>>>() {}.getType();
                Collection<List<String>> list = new Gson().fromJson(jsonArray, collectionType);
                resultTabletRow.addAll(list);
            }
        } else {
            WorkbookTableRowCollectionPage workbookTableRowCollectionPage = graphClient.me().drive().items(driveItemId).workbook().worksheets(workSheetId)
                    .tables(tableId).rows().buildRequest().get();
            if (workbookTableRowCollectionPage != null) {
                List<WorkbookTableRow> tableRows = workbookTableRowCollectionPage.getCurrentPage();
                for (WorkbookTableRow tableRow : tableRows) {
                    if (tableRow.values != null) {
                        JsonElement values = tableRow.values.getAsJsonArray().get(0);
                        Type collectionType = new TypeToken<List<String>>() {}.getType();
                        List<String> list = new Gson().fromJson(values, collectionType);
                        resultTabletRow.add(list);
                    }
                }
            }
        }
        return resultTabletRow;
    }
}
