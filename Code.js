function doGet(e) {
    var requestParams = {};
    // Initiate HTML UI component 
    var pageParam = e.parameters['page'];
    requestParams = JSON.parse(pageParam);
    // const keys = Object.keys(e.parameters);
    var page = requestParams.page;
    var htmlTemplate = HtmlService.createTemplateFromFile(page);

    htmlTemplate.requestParams = requestParams;
    var output = htmlTemplate.evaluate();
    // set title by page
    var title;
    switch (page) {
        case 'manage':
            title = '投票管理';
            break;
        case 'form':
            title = '投票詳細内容';
            break;
        case 'previewForm':
            title = '投票詳細内容';
            break;
        case 'qrCode':
            title = 'QRコード';
            break;
        case 'statistic':
            title = '結果の統計';
            break;
        case 'thankyou':
            title = '投票成功';
            break;
        default:
            title = '投票';
            break;
    }
    
    output.setTitle(title);
    output.setSandboxMode(HtmlService.SandboxMode.NATIVE);
    output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    ui = output.addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1, user-scalable=no');
    return ui;
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}