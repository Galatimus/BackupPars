# -*- coding: utf-8 -*-
#!/usr/bin/python
conv =[(u'Абакан',u'Республика Хакасия'),
    (u'Дубовая Роща',u'Московская область'),
    (u'Мосрентген',u'Москва'),
    (u'Федино',u'Московская область'),
    (u'Электроугли',u'Московская область'),
    (u'Яхрома',u'Московская область'),
    (u'Абинск',u'Краснодарский край'),
    (u'Индустриальный',u'Краснодарский край'),
    (u'Азнакаево',u'Татарстан'),
    (u'Азов',u'Ростовская область'),
    (u'Аксай',u'Ростовская область'),
    (u'Алатырь',u'Чувашия'),
    (u'Алейск',u'Алтайский край'),
    (u'Александров',u'Владимирская область'),
    (u'Алексеевка',u'Белгородская область'),
    (u'Алексин',u'Тульская область'),
    (u'Алупка',u'Крым'),
    (u'Алушта',u'Крым'),
    (u'Альметьевск',u'Татарстан'),
    (u'Амурск',u'Хабаровский край'),
    (u'Анадырь',u'Чукотский автономный округ'),
    (u'Анапа',u'Краснодарский край'),
    (u'Ангарск',u'Иркутская область'),
    (u'Анива',u'Сахалинская область'),
    (u'Апатиты',u'Мурманская область'),
    (u'Апшеронск',u'Краснодарский край'),
    (u'Арамиль',u'Свердловская область'),
    (u'Ардатов',u'Мордовия'),
    (u'Арзамас',u'Нижегородская область'),
    (u'Армавир',u'Краснодарский край'),
    (u'Арсеньев',u'Приморский край'),
    (u'Артем',u'Приморский край'),
    (u'Архангельск',u'Архангельская область'),
    (u'Астрахань',u'Астраханская область'),
    (u'Ахтубинск',u'Астраханская область'),
    (u'Ачинск',u'Красноярский край'),
    (u'Бакал',u'Челябинская область'),
    (u'Баксан',u'Кабардино-Балкария'),
    (u'Балаково',u'Саратовская область'),
    (u'Балашиха',u'Московская область'),
    (u'Балашов',u'Саратовская область'),
    (u'Барнаул',u'Алтайский край'),
    (u'Барыш',u'Ульяновская область'),
    (u'Батайск',u'Ростовская область'),
    (u'Бахчисарай',u'Крым'),
    (u'Белая Калитва',u'Ростовская область'),
    (u'Белгород',u'Белгородская область'),
    (u'Белев',u'Тульская область'),
    (u'Белово',u'Кемеровская область'),
    (u'Белогорск',u'Амурская область'),
    (u'Белокуриха',u'Алтайский край'),
    (u'Белорецк',u'Башкортостан'),
    (u'Белореченск',u'Краснодарский край'),
    (u'Бердск',u'Новосибирская область'),
    (u'Бийск',u'Алтайский край'),
    (u'Бикин',u'Хабаровский край'),
    (u'Биробиджан',u'Еврейская автономная область'),
    (u'Благовещенск',u'Амурская область'),
    (u'Богданович',u'Свердловская область'),
    (u'Богородск',u'Нижегородская область'),
    (u'Богучаны',u'Красноярский край'),
    (u'Богучар',u'Воронежская область'),
    (u'Большой Камень',u'Приморский край'),
    (u'Бор',u'Нижегородская область'),
    (u'Борисоглебск',u'Воронежская область'),
    (u'Боровичи',u'Новгородская область'),
    (u'Боровск',u'Калужская область'),
    (u'Братск',u'Иркутская область'),
	(u'Березники',u'Пермский край'),
    (u'Бронницы',u'Московская область'),
    (u'Брянск',u'Брянская область'),
    (u'Бугульма',u'Татарстан'),
    (u'Бугуруслан',u'Оренбургская область'),
    (u'Бузулук',u'Оренбургская область'),
    (u'Ванино',u'Хабаровский край'),
    (u'Великие Луки',u'Псковская область'),
    (u'Великий Новгород',u'Новгородская область'),
    (u'Вельск',u'Архангельская область'),
    (u'Венев',u'Тульская область'),
    (u'Верещагино',u'Пермский край'),
    (u'Верея',u'Московская область'),
    (u'Верхняя Пышма',u'Свердловская область'),
    (u'Видное',u'Московская область'),
    (u'Вилючинск',u'Камчптский край'),
    (u'Вичуга',u'Ивановская область'),
    (u'Владивосток',u'Приморский край'),
    (u'Владикавказ',u'Республика Северная Осетия-Алания'),
    (u'Владимир',u'Владимирская область'),
    (u'Волгоград',u'Волгоградская область'),
    (u'Волгодонск',u'Ростовская область'),
    (u'Волжский',u'Волгоградская область'),
    (u'Вологда',u'Вологодская область'),
    (u'Волоколамск',u'Московская область'),
    (u'Волхов',u'Ленинградская область'),
    (u'Воронеж',u'Воронежская область'),
    (u'Воткинск',u'Удмуртская республика'),
    (u'Всеволожск',u'Ленинградская область'),
    (u'Выборг',u'Ленинградская область'),
    (u'Выкса',u'Нижегородская область'),
    (u'Вяземский',u'Хабаровский край'),
    (u'Вязники',u'Владимирская область'),
    (u'Вятские Поляны',u'Кировская область'),
    (u'Гай',u'Оренбургская область'),
    (u'Геленджик',u'Краснодарский край'),
    (u'Георгиевск',u'Ставропольский край'),
    (u'Горно-Алтайск',u'Республика Алтай'),
    (u'Городище',u'Московская область'),
    (u'Грозный',u'Чеченская Республика'),
    (u'Гусь-Хрустальный',u'Московская область'),
    (u'Воскресенск',u'Московская область'),
    (u'Дальнегорск',u'Приморский край'),
    (u'Дальнереченск',u'Приморский край'),
    (u'Дзержинск',u'Нижегородская область'),
    (u'Дзержинский',u'Московская область'),
    (u'Дивногорск',u'Красноярский край'),
    (u'Димитровград',u'Ульяновская область'),
    (u'Дмитров',u'Московская область'),
    (u'Долгопрудный',u'Московская область'),
    (u'Домодедово',u'Московская область'),
    (u'Дюртюли',u'Башкортостан'),
    (u'Евпатория',u'Крым'),
    (u'Егорьевск',u'Московская область'),
    (u'Ейск',u'Краснодарский край'),
    (u'Екатеринбург',u'Свердловская область'),
    (u'Елец',u'Липецкая область'),
    (u'Елизово',u'Камчптский край'),
    (u'Енисейск',u'Красноярский край'),
    (u'Ершов',u'Саратовская область'),
    (u'Ессентуки',u'Ставропольский край'),
    (u'Ефремов',u'Тульская область'),
    (u'Железноводск',u'Ставропольский край'),
    (u'Железнодорожный',u'Московская область'),
    (u'Жуковский',u'Московская область'),
    (u'Завитинск',u'Амурская область'),
    (u'Задонск',u'Липецкая область'),
    (u'Заозерный',u'Красноярский край'),
    (u'Заринск',u'Алтайский край'),
    (u'Звенигород',u'Московская область'),
    (u'Зверево',u'Ростовская область'),
    (u'Зеленогорск',u'Красноярский край'),
    (u'Зеленоград',u'Москва'),
    (u'Зеленокумск',u'Ставропольский край'),
    (u'Зерноград',u'Ростовская область'),
    (u'Зерноград',u'Ростовская область'),
    (u'Зея',u'Амурская область'),
    (u'Златоуст',u'Челябинская область'),
    (u'Змеиногорск',u'Алтайский край'),
    (u'Знаменск',u'Астраханская область'),
    (u'Зубцов',u'Тверская область'),
    (u'Зуевка',u'Кировская область'),
    (u'Иваново',u'Ивановская область'),
    (u'Ивантеевка',u'Московская область'),
    (u'Ижевск',u'Удмуртская республика'),
    (u'Ипатово',u'Ставропольский край'),
    (u'Иркутск',u'Иркутская область'),
    (u'Искитим',u'Новосибирская область'),
    (u'Истра',u'Московская область'),
    (u'Ишим',u'Тюменская область'),
    (u'Йошкар-Ола',u'Республика Марий Эл'),
    (u'Казань',u'Республика Татарстан'),
    (u'Калач-на-Дону',u'Волгоградская область'),
    (u'Калачинск',u'Омская область'),
    (u'Калининград',u'Калининградская область'),
    (u'Калуга',u'Калужская область'),
    (u'Калязин',u'Тверская область'),
    (u'Каменка',u'Воронежская область'),
    (u'Каменск-Уральский',u'Свердловская область'),
    (u'Каменск-Шахтинский',u'Ростовская область'),
    (u'Камень-на-Оби',u'Алтайский край'),
    (u'Камызяк',u'Астраханская область'),
    (u'Камышин',u'Волгоградская область'),
    (u'Канск',u'Красноярский край'),
    (u'Карасук',u'Новосибирская область'),
    (u'Карачаевск',u'Карачаево-Черкесия'),
    (u'Каспийск',u'Республика Дагестан'),
    (u'Кашин',u'Тверская область'),
    (u'Кашира',u'Московская область'),
    (u'Кемерово',u'Кемеровская область'),
    (u'Керчь',u'Крым'),
    (u'Кимры',u'Тверская область'),
    (u'Кингисепп',u'Ленинградская область'),
    (u'Кинель',u'Самарская область'),
    (u'Кинешма',u'Ивановская область'),
    (u'Киреевск',u'Тульская область'),
    (u'Кириши',u'Ленинградская область'),
    (u'Киров',u'Кировская область'),
    (u'Кирово-Чепецк',u'Кировская область'),
    (u'Киселевск',u'Кемеровская область'),
    (u'Кисловодск',u'Ставропольский край'),
    (u'Климовск',u'Московская область'),
    (u'Клин',u'Московская область'),
    (u'Козьмодемьянск',u'Марий Эл'),
    (u'Коломна',u'Московская область'),
    (u'Колпино',u'Санкт-Петербург'),
    (u'Комсомольск-на-Амуре',u'Хабаровский край'),
    (u'Кондрово',u'Калужская область'),
    (u'Константиновск',u'Ростовская область'),
    (u'Копейск',u'Челябинская область'),
    (u'Кореновск',u'Краснодарский край'),
    (u'Королев',u'Московская область'),
    (u'Королёв',u'Московская область'),
    (u'Кострома',u'Костромская область'),
    (u'Котельники',u'Московская область'),
    (u'Котлас',u'Архангельская область'),
    (u'Красногорск',u'Московская область'),
    (u'Краснодар',u'Краснодарский край'),
    (u'Краснокаменск',u'Забайкальский край'),
    (u'Краснослободск',u'Волгоградская область'),
    (u'Красноуфимск',u'Свердловская область'),
    (u'Красноярск',u'Красноярский край'),
    (u'Красный Сулин',u'Ростовская область'),
    (u'Кропоткин',u'Краснодарский край'),
    (u'Крымск',u'Краснодарский край'),
    (u'Кубинка',u'Московская область'),
    (u'Кузнецк',u'Пензенская область'),
    (u'Кумертау',u'Башкортостан'),
    (u'Курган',u'Курганская область'),
    (u'Курск',u'Курская область'),
    (u'Кызыл',u'Тыва'),
    (u'Лабинск',u'Краснодарский край'),
    (u'Лазо',u'Камчатский край'),
    (u'Лангепас',u'Ханты-Мансийский автономный округ — Югра'),
    (u'Ленинск',u'Пензенская областьий'),
    (u'Лесозаводск',u'Приморский край'),
    (u'Ликино-Дулево',u'Московская область'),
    (u'Липецк',u'Липецкая область'),
    (u'Лихославль',u'Тверская область'),
    (u'Лобня',u'Московская область'),
    (u'Ломоносов',u'Ленинградская область'),
    (u'Лучегорск',u'Приморский край'),
    (u'Лучегорск',u'Приморский край'),
    (u'Лыткарино',u'Московская область'),
    (u'Люберцы',u'Московская область'),
    (u'Малаховка',u'Московская область'),
    (u'Апаринки',u'Московская область'),
    (u'Ашитково',u'Московская область'),
    (u'Львовский',u'Московская область'),
    (u'Магадан',u'Магаданская область'),
    (u'Магас',u'Ингушетия'),
    (u'Магнитогорск',u'Челябинская область'),
    (u'Майкоп',u'Адыгея'),
    (u'Малоярославец',u'Калужская область'),
    (u'Мантурово',u'Костромская область'),
    (u'Мариинск',u'Кемеровская область'),
    (u'Махачкала',u'Республика Дагестан'),
    (u'Междуреченск',u'Кемеровская область'),
    (u'Миасс',u'Челябинская область'),
    (u'Минеральные Воды',u'Ставропольский край'),
    (u'Минеральные воды',u'Ставропольский край'),
    (u'Минусинск',u'Красноярский край'),
    (u'Мирный',u'Приморский край'),
    (u'Михайлов',u'Рязанская область'),
    (u'Михайловка',u'Волгоградская область'),
    (u'Михайловск',u'Ставропольский край'),
    (u'Москвы',u'Москва'),
    (u'Мурманск',u'Мурманская область'),
    (u'Муром',u'Владимирская область'),
    (u'Мценск',u'Орловская область'),
    (u'Мыски',u'Кемеровская область'),
    (u'Мытищи',u'Московская область'),
    (u'Набережные Челны',u'Республика Татарстан'),
    (u'Надым',u'Ямало-Ненецкий автономный округ'),
    (u'Нальчик',u'Кабардино-Балкария'),
    (u'Новая Адыгея',u'Адыгея'),
    (u'Наро-Фоминск',u'Московская область'),
    (u'Нахабино',u'Московская область'),
    (u'Находка',u'Приморский край'),
    (u'Нерюнгри',u'Республика Саха (Якутия)'),
    (u'Нефтекамск',u'Башкортостан'),
    (u'Нефтеюганск',u'Ханты-Мансийский автономный округ—Югра'),
    (u'Нижневартовск',u'Ханты-Мансийский автономный округ—Югра'),
    (u'Нижнекамск',u'Республика Татарстан'),
    (u'Нижний Новгород',u'Нижегородская область'),
    (u'Нижний Тагил',u'Свердловская область'),
    (u'Николаевск-на-Амуре',u'Хабаровский край'),
    (u'Никольское',u'Сахалинская область'),
    (u'Новоалтайск',u'Алтайский край'),
    (u'Новозыбков',u'Брянская область'),
    (u'Новокузнецк',u'Кемеровская область'),
    (u'Новомосковск',u'Тульская область'),
    (u'Новороссийск',u'Краснодарский край'),
    (u'Новосибирск',u'Новосибирская область'),
    (u'Новотроицк',u'Оренбургская область'),
    (u'Новочебоксарск',u'Чувашия'),
    (u'Новочеркасск',u'Ростовская область'),
    (u'Новый Уренгой',u'Ямало-Ненецкий автономный округ'),
    (u'Ногинск',u'Московская область'),
    (u'Норильск',u'Красноярский край'),
    (u'Ноябрьск',u'Ямало-Ненецкий автономный округ'),
    (u'Нягань',u'Ханты-Мансийский автономный округ — Югра'),
    (u'Облучье',u'Еврейская автономная область'),
    (u'Обнинск',u'Калужская область'),
    (u'Одинцово',u'Московская область'),
    (u'Ожерелье',u'Московская область'),
    (u'Октябрьский',u'Башкортостан'),
    (u'Омск',u'Омская область'),
    (u'Орел',u'Орловская область'),
    (u'Оренбург',u'Оренбургская область'),
    (u'Орехово-Зуево',u'Московская область'),
    (u'Орск',u'Оренбургская область'),
    (u'Остров',u'Псковская область'),
    (u'Павлово',u'Нижегородская область'),
    (u'Павловск',u'Воронежская область'),
    (u'Партизанск',u'Приморский край'),
    (u'Пенза',u'Пензенская область'),
    (u'Переславль-Залесский',u'Ярославская область'),
    (u'Пермь',u'Пермский край'),
    (u'Пушкин',u'Санкт-Петербург'),
    (u'Парголово',u'Санкт-Петербург'),
    (u'Петрозаводск',u'Республика Карелия'),
    (u'Петропавловск-Камчатский',u'Камчатский край'),
    (u'Пласт',u'Челябинская область'),
    (u'Плес',u'Волгоградская область'),
    (u'Подольск',u'Московская область'),
    (u'Порхов',u'Псковская область'),
    (u'Починок',u'Смоленская область'),
    (u'Преображение',u'Приморский'),
    (u'Приморский',u'Приморский край'),
    (u'Приморско-Ахтарск',u'Краснодарский край'),
    (u'Приозерск',u'Ленинградская область'),
    (u'Тосно',u'Ленинградская область'),
    (u'Волосово',u'Ленинградская область'),
    (u'Прокопьевск',u'Кемеровская область'),
    (u'Пролетарск',u'Ростовская область'),
    (u'Псков',u'Псковская область'),
    (u'Пушкино',u'Московская область'),
    (u'Хорлово',u'Московская область'),
    (u'Белоозерский пгт',u'Московская область'),
    (u'Томилино',u'Московская область'),
    (u'Электрогорск',u'Московская область'),
    (u'Старая Купавна',u'Московская область'),
    (u'Северный',u'Московская область'),
    (u'Витаминкомбинат',u'Краснодарский край'),
    (u'Динская',u'Краснодарский край'),
    (u'Елизаветинская',u'Ростовская область'),
    (u'Пятигорск',u'Ставропольский край'),
    (u'Радужный',u'Ханты-Мансийский автономный округ — Югра'),
    (u'Райчихинск',u'Амурская область'),
    (u'Раменское',u'Московская область'),
    (u'Рассказово',u'Тамбовская область'),
    (u'Ревда',u'Свердловская область'),
    (u'Реутов',u'Московская область'),
    (u'Ржев',u'Тверская область'),
    (u'Россошь',u'Воронежская область'),
    (u'Ростов-на-Дону',u'Ростовская область'),
    (u'Рубцовск',u'Алтайский край'),
    (u'Рудня',u'Смоленская область'),
    (u'Рузаевка',u'Мордовия'),
    (u'Ряжск',u'Рязанская область'),
    (u'Рязань',u'Рязанская область'),
    (u'Саки',u'Крым'),
    (u'Салават',u'Республика Башкортостан'),
    (u'Салехард',u'Ямало-Ненецкий автономный округ'),
    (u'Сальск',u'Ростовская область'),
    (u'Самара',u'Самарская область'),
    (u'Саранск',u'Республика Мордовия'),
    (u'Саратов',u'Саратовская область'),
    (u'Саров',u'Нижегородская область'),
    (u'Саяногорск',u'Хакасия'),
    (u'Саянск',u'Иркутская область'),
    (u'Светлоград',u'Ставропольский край'),
    (u'Свободный',u'Амурская область'),
    (u'Северодвинск',u'Архангельская область'),
    (u'Североуральск',u'Свердловская область'),
    (u'Северск',u'Томская область'),
    (u'Семенов',u'Нижегородская область'),
    (u'Сенгилей',u'Ульяновская область'),
    (u'Сергиев Посад',u'Московская область'),
    (u'Серпухов',u'Московская область'),
    (u'Сертолово',u'Ленинградская область'),
    (u'Серышево-2',u'Амурская область'),
    (u'Сибай',u'Башкортостан'),
    (u'Симеиз',u'Крым'),
    (u'Симферополь',u'Крым'),
    (u'Сковородино',u'Амурская область'),
    (u'Славянск-на-Кубани',u'Краснодарский край'),
    (u'Смирных',u'Сахалинская область'),
    (u'Смоленск',u'Смоленская область'),
    (u'Собинка',u'Владимирская область'),
    (u'Советская Гавань',u'Хабаровский край'),
    (u'Соликамск',u'Пермский край'),
    (u'Солнечногорск',u'Московская область'),
    (u'Сосновоборск',u'Красноярский край'),
    (u'Сосновый Бор',u'Ленинградская область'),
    (u'Сочи',u'Краснодарский край'),
    (u'Спасск-Дальний',u'Приморский край'),
    (u'Ставрополь',u'Ставропольский край'),
    (u'Старый Крым',u'Крым'),
    (u'Старый Оскол',u'Белгородская область'),
    (u'Стерлитамак',u'Республика Башкортостан'),
    (u'Строитель',u'Белгородская область'),
    (u'Ступино',u'Московская область'),
    (u'Судак',u'Крым'),
    (u'Суджа',u'Курская область'),
    (u'Суздаль',u'Владимирская область'),
    (u'Сургут',u'Ханты-Мансийский автономный округ—Югра'),
    (u'Сухиничи',u'Калужская область'),
    (u'Сухой Лог',u'Свердловская область'),
    (u'Сходня',u'Московская область'),
    (u'Сызрань',u'Самарская область'),
    (u'Сыктывкар',u'Республика Коми'),
    (u'Сысерть',u'Свердловская область'),
    (u'Тавричанка',u'Приморский край'),
    (u'Таганрог',u'Ростовская область'),
    (u'Талдом',u'Московская область'),
    (u'Фрязино',u'Московская область'),
    (u'Серебряные Пруды',u'Московская область'),
    (u'Тамбов',u'Тамбовская область'),
    (u'Тверь',u'Тверская область'),
    (u'Тейково',u'Ивановская область'),
    (u'Темрюк',u'Краснодарский край'),
    (u'Тимашевск',u'Краснодарский край'),
    (u'Тихвин',u'Ленинградская область'),
    (u'Гатчина',u'Ленинградская область'),
    (u'Троицк',u'Москва'),
    (u'Щербинка',u'Москва'),
    (u'Тобольск',u'Тюменская область'),
    (u'Тольятти',u'Самарская область'),
    (u'Томмот',u'Якутия'),
    (u'Бараба-Кулики',u'Свердловская область'),
    (u'Старая Русса',u'Новгородская область'),
    (u'Томск',u'Томская область'),
    (u'Торопец',u'Тверская область'),
    (u'Туапсе',u'Краснодарский край'),
    (u'Тула',u'Тульская область'),
    (u'Тулун',u'Иркутская область'),
    (u'Тында',u'Амурская область'),
    (u'Тюмень',u'Тюменская область'),
    (u'Улан-Удэ',u'Республика Бурятия'),
    (u'Ульяновск',u'Ульяновская область'),
    (u'Урай',u'Ханты-Мансийский автономный округ — Югра'),
    (u'Урюпинск',u'Волгоградская область'),
    (u'Усинск',u'Республика Коми'),
    (u'Усмань',u'Липецкая область'),
    (u'Усолье-Сибирское',u'Иркутская область'),
    (u'Уссурийск',u'Приморский край'),
    (u'Усть-Илимск',u'Иркутская область'),
    (u'Уфа',u'Республика Башкортастан'),
    (u'Ухта',u'Республика Коми'),
    (u'Феодосия',u'Крым'),
    (u'Яблоновский',u'Адыгея'),
    (u'Фокино',u'Приморский край'),
    (u'Хабаровск',u'Хабаровский край'),
    (u'Ханты-Мансийск',u'Ханты-Мансийский автономный округ — Югра'),
    (u'Хасавюрт',u'Дагестан'),
    (u'Химки',u'Московская область'),
    (u'Хороль',u'Приморский край'),
    (u'Хотьково',u'Московская область'),
    (u'Чайковский',u'Пермский край'),
    (u'Чебаркуль',u'Челябинская область'),
    (u'Чебоксары',u'Чувашская Республика'),
    (u'Челябинск',u'Челябинская область'),
    (u'Черемхово',u'Иркутская область'),
    (u'Черепаново',u'Новосибирская область'),
    (u'Череповец',u'Вологодская область'),
    (u'Черкесск',u'Карачаево-Черкесская Республика'),
    (u'Черноморское',u'Крым'),
    (u'Чехов',u'Московская область'),
    (u'Чита',u'Забайкальскй край'),
    (u'Энем',u'Адыгея'),
    (u'Рассказовка',u'Москва'),
    (u'Шатура',u'Московская область'),
    (u'Шахты',u'Ростовская область'),
    (u'Шелехов',u'Иркутская область'),
    (u'Шимановск',u'Амурская область'),
    (u'Шлиссельбург',u'Ленинградская область'),
    (u'Луга',u'Ленинградская область'),
    (u'Щекино',u'Тульская область'),
    (u'Щелково',u'Московская область'),
    (u'Новотитаровская',u'Краснодарский край'),
    (u'Лучинское',u'Московская область'),
    (u'Шереметьево',u'Московская область'),
    (u'Электросталь',u'Московская область'),
    (u'Элиста',u'Калмыкия'),
    (u'Энгельс',u'Саратовская область'),
    (u'Югорск',u'Ханты-Мансийский автономный округ—Югра'),
    (u'Южно-Сахалинск',u'Свердловская область'),
    (u'Южноуральск',u'Челябинская область'),
    (u'Юрга',u'Кемеровская область'),
    (u'Юрьев-Польский',u'Владимирская область'),
    (u'Юхнов',u'Калужская область'),
    (u'Якутск',u'Саха'),
    (u'Ялта',u'Крым'),
    (u'Ялуторовск',u'Тюменская область'),
    (u'Яровое',u'Алтайский край'),
    (u'Ярославль',u'Ярославская область'),
    (u'Ясногорск',u'Тульская область')]