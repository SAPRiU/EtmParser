Вся работа с API ETM выполняется в файле EtmParserViewModel.cs.
EtmParserViewModel.cs содержит 2 основных класса: EtmParserViewModel и ApiRequester.
EtmParserViewModel содержит в себе логику по чтению и заполнению Excel таблицы и редактировании конфиг файла приложения, в котором хранятся учетные данные. EtmParserViewModel использует экземпляр класса ApiRequester для получения данных с API.
ApiRequester реализует непосредственное обращение к API. Имеет 3 основных метода: AuthorizationRequest, GoodsRequest и RemainRequest.
AuthorizationRequest - запрос для авторизации и получения токена.
GoodsRequest - запрос для получения характеристик товаров.
RemainRequest - запрос по остаткам товара.
