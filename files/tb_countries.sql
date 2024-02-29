/****** Object:  Table [dbo].[tb_countries]    Script Date: 06/22/2017 11:42:13 ******/
CREATE TABLE [dbo].[tb_countries](
	[seq] [int] NULL,
	[name] [text] NOT NULL,
	[code] [text] NULL,
	[currency] [text] NULL,
	[population] [text] NULL,
	[capital] [text] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]


INSERT INTO tb_countries (code, name, currency, population, capital) VALUES
('AD', 'Andorra', 'EUR', '84000', 'Andorra la Vella'),
('AE', 'United Arab Emirates', 'AED', '4975593', 'Abu Dhabi'),
('AF', 'Afghanistan', 'AFN', '29121286', 'Kabul'),
('AG', 'Antigua and Barbuda', 'XCD', '86754', 'St. Johns'),
('AI', 'Anguilla', 'XCD', '13254', 'The Valley'),
('AL', 'Albania', 'ALL', '2986952', 'Tirana'),
('AM', 'Armenia', 'AMD', '2968000', 'Yerevan'),
('AO', 'Angola', 'AOA', '13068161', 'Luanda'),
('AQ', 'Antarctica', '', '0', ''),
('AR', 'Argentina', 'ARS', '41343201', 'Buenos Aires'),
('AS', 'American Samoa', 'USD', '57881', 'Pago Pago'),
('AT', 'Austria', 'EUR', '8205000', 'Vienna'),
('AU', 'Australia', 'AUD', '21515754', 'Canberra'),
('AW', 'Aruba', 'AWG', '71566', 'Oranjestad'),
('AX', 'Åland', 'EUR', '26711', 'Mariehamn'),
('AZ', 'Azerbaijan', 'AZN', '8303512', 'Baku'),
('BA', 'Bosnia and Herzegovina', 'BAM', '4590000', 'Sarajevo'),
('BB', 'Barbados', 'BBD', '285653', 'Bridgetown'),
('BD', 'Bangladesh', 'BDT', '156118464', 'Dhaka'),
('BE', 'Belgium', 'EUR', '10403000', 'Brussels'),
('BF', 'Burkina Faso', 'XOF', '16241811', 'Ouagadougou'),
('BG', 'Bulgaria', 'BGN', '7148785', 'Sofia'),
('BH', 'Bahrain', 'BHD', '738004', 'Manama'),
('BI', 'Burundi', 'BIF', '9863117', 'Bujumbura'),
('BJ', 'Benin', 'XOF', '9056010', 'Porto-Novo'),
('BL', 'Saint Barthélemy', 'EUR', '8450', 'Gustavia'),
('BM', 'Bermuda', 'BMD', '65365', 'Hamilton'),
('BN', 'Brunei', 'BND', '395027', 'Bandar Seri Begawan'),
('BO', 'Bolivia', 'BOB', '9947418', 'Sucre'),
('BQ', 'Bonaire', 'USD', '18012', 'Kralendijk'),
('BR', 'Brazil', 'BRL', '201103330', 'Brasília'),
('BS', 'Bahamas', 'BSD', '301790', 'Nassau'),
('BT', 'Bhutan', 'BTN', '699847', 'Thimphu'),
('BV', 'Bouvet Island', 'NOK', '0', ''),
('BW', 'Botswana', 'BWP', '2029307', 'Gaborone'),
('BY', 'Belarus', 'BYN', '9685000', 'Minsk'),
('BZ', 'Belize', 'BZD', '314522', 'Belmopan'),
('CA', 'Canada', 'CAD', '33679000', 'Ottawa'),
('CC', 'Cocos [Keeling] Islands', 'AUD', '628', 'West Island'),
('CD', 'Democratic Republic of the Congo', 'CDF', '70916439', 'Kinshasa'),
('CF', 'Central African Republic', 'XAF', '4844927', 'Bangui'),
('CG', 'Republic of the Congo', 'XAF', '3039126', 'Brazzaville'),
('CH', 'Switzerland', 'CHF', '7581000', 'Bern'),
('CI', 'Ivory Coast', 'XOF', '21058798', 'Yamoussoukro'),
('CK', 'Cook Islands', 'NZD', '21388', 'Avarua'),
('CL', 'Chile', 'CLP', '16746491', 'Santiago'),
('CM', 'Cameroon', 'XAF', '19294149', 'Yaoundé'),
('CN', 'China', 'CNY', '1330044000', 'Beijing'),
('CO', 'Colombia', 'COP', '47790000', 'Bogotá'),
('CR', 'Costa Rica', 'CRC', '4516220', 'San José'),
('CU', 'Cuba', 'CUP', '11423000', 'Havana'),
('CV', 'Cape Verde', 'CVE', '508659', 'Praia'),
('CW', 'Curacao', 'ANG', '141766', 'Willemstad'),
('CX', 'Christmas Island', 'AUD', '1500', 'Flying Fish Cove'),
('CY', 'Cyprus', 'EUR', '1102677', 'Nicosia'),
('CZ', 'Czechia', 'CZK', '10476000', 'Prague'),
('DE', 'Germany', 'EUR', '81802257', 'Berlin'),
('DJ', 'Djibouti', 'DJF', '740528', 'Djibouti'),
('DK', 'Denmark', 'DKK', '5484000', 'Copenhagen'),
('DM', 'Dominica', 'XCD', '72813', 'Roseau'),
('DO', 'Dominican Republic', 'DOP', '9823821', 'Santo Domingo'),
('DZ', 'Algeria', 'DZD', '34586184', 'Algiers'),
('EC', 'Ecuador', 'USD', '14790608', 'Quito'),
('EE', 'Estonia', 'EUR', '1291170', 'Tallinn'),
('EG', 'Egypt', 'EGP', '80471869', 'Cairo'),
('EH', 'Western Sahara', 'MAD', '273008', 'Laâyoune / El Aaiún'),
('ER', 'Eritrea', 'ERN', '5792984', 'Asmara'),
('ES', 'Spain', 'EUR', '46505963', 'Madrid'),
('ET', 'Ethiopia', 'ETB', '88013491', 'Addis Ababa'),
('FI', 'Finland', 'EUR', '5244000', 'Helsinki'),
('FJ', 'Fiji', 'FJD', '875983', 'Suva'),
('FK', 'Falkland Islands', 'FKP', '2638', 'Stanley'),
('FM', 'Micronesia', 'USD', '107708', 'Palikir'),
('FO', 'Faroe Islands', 'DKK', '48228', 'Tórshavn'),
('FR', 'France', 'EUR', '64768389', 'Paris'),
('GA', 'Gabon', 'XAF', '1545255', 'Libreville'),
('GB', 'United Kingdom', 'GBP', '62348447', 'London'),
('GD', 'Grenada', 'XCD', '107818', 'St. Georges'),
('GE', 'Georgia', 'GEL', '4630000', 'Tbilisi'),
('GF', 'French Guiana', 'EUR', '195506', 'Cayenne'),
('GG', 'Guernsey', 'GBP', '65228', 'St Peter Port'),
('GH', 'Ghana', 'GHS', '24339838', 'Accra'),
('GI', 'Gibraltar', 'GIP', '27884', 'Gibraltar'),
('GL', 'Greenland', 'DKK', '56375', 'Nuuk'),
('GM', 'Gambia', 'GMD', '1593256', 'Bathurst'),
('GN', 'Guinea', 'GNF', '10324025', 'Conakry'),
('GP', 'Guadeloupe', 'EUR', '443000', 'Basse-Terre'),
('GQ', 'Equatorial Guinea', 'XAF', '1014999', 'Malabo'),
('GR', 'Greece', 'EUR', '11000000', 'Athens'),
('GS', 'South Georgia and the South Sandwich Islands', 'GBP', '30', 'Grytviken'),
('GT', 'Guatemala', 'GTQ', '13550440', 'Guatemala City'),
('GU', 'Guam', 'USD', '159358', 'Hagåtña'),
('GW', 'Guinea-Bissau', 'XOF', '1565126', 'Bissau'),
('GY', 'Guyana', 'GYD', '748486', 'Georgetown'),
('HK', 'Hong Kong', 'HKD', '6898686', 'Hong Kong'),
('HM', 'Heard Island and McDonald Islands', 'AUD', '0', ''),
('HN', 'Honduras', 'HNL', '7989415', 'Tegucigalpa'),
('HR', 'Croatia', 'HRK', '4284889', 'Zagreb'),
('HT', 'Haiti', 'HTG', '9648924', 'Port-au-Prince'),
('HU', 'Hungary', 'HUF', '9982000', 'Budapest'),
('ID', 'Indonesia', 'IDR', '242968342', 'Jakarta'),
('IE', 'Ireland', 'EUR', '4622917', 'Dublin'),
('IL', 'Israel', 'ILS', '7353985', ''),
('IM', 'Isle of Man', 'GBP', '75049', 'Douglas'),
('IN', 'India', 'INR', '1173108018', 'New Delhi'),
('IO', 'British Indian Ocean Territory', 'USD', '4000', ''),
('IQ', 'Iraq', 'IQD', '29671605', 'Baghdad'),
('IR', 'Iran', 'IRR', '76923300', 'Tehran'),
('IS', 'Iceland', 'ISK', '308910', 'Reykjavik'),
('IT', 'Italy', 'EUR', '60340328', 'Rome'),
('JE', 'Jersey', 'GBP', '90812', 'Saint Helier'),
('JM', 'Jamaica', 'JMD', '2847232', 'Kingston'),
('JO', 'Jordan', 'JOD', '6407085', 'Amman'),
('JP', 'Japan', 'JPY', '127288000', 'Tokyo'),
('KE', 'Kenya', 'KES', '40046566', 'Nairobi'),
('KG', 'Kyrgyzstan', 'KGS', '5776500', 'Bishkek'),
('KH', 'Cambodia', 'KHR', '14453680', 'Phnom Penh'),
('KI', 'Kiribati', 'AUD', '92533', 'Tarawa'),
('KM', 'Comoros', 'KMF', '773407', 'Moroni'),
('KN', 'Saint Kitts and Nevis', 'XCD', '51134', 'Basseterre'),
('KP', 'North Korea', 'KPW', '22912177', 'Pyongyang'),
('KR', 'South Korea', 'KRW', '48422644', 'Seoul'),
('KW', 'Kuwait', 'KWD', '2789132', 'Kuwait City'),
('KY', 'Cayman Islands', 'KYD', '44270', 'George Town'),
('KZ', 'Kazakhstan', 'KZT', '15340000', 'Astana'),
('LA', 'Laos', 'LAK', '6368162', 'Vientiane'),
('LB', 'Lebanon', 'LBP', '4125247', 'Beirut'),
('LC', 'Saint Lucia', 'XCD', '160922', 'Castries'),
('LI', 'Liechtenstein', 'CHF', '35000', 'Vaduz'),
('LK', 'Sri Lanka', 'LKR', '21513990', 'Colombo'),
('LR', 'Liberia', 'LRD', '3685076', 'Monrovia'),
('LS', 'Lesotho', 'LSL', '1919552', 'Maseru'),
('LT', 'Lithuania', 'EUR', '2944459', 'Vilnius'),
('LU', 'Luxembourg', 'EUR', '497538', 'Luxembourg'),
('LV', 'Latvia', 'EUR', '2217969', 'Riga'),
('LY', 'Libya', 'LYD', '6461454', 'Tripoli'),
('MA', 'Morocco', 'MAD', '33848242', 'Rabat'),
('MC', 'Monaco', 'EUR', '32965', 'Monaco'),
('MD', 'Moldova', 'MDL', '4324000', 'Chişinău'),
('ME', 'Montenegro', 'EUR', '666730', 'Podgorica'),
('MF', 'Saint Martin', 'EUR', '35925', 'Marigot'),
('MG', 'Madagascar', 'MGA', '21281844', 'Antananarivo'),
('MH', 'Marshall Islands', 'USD', '65859', 'Majuro'),
('MK', 'Macedonia', 'MKD', '2062294', 'Skopje'),
('ML', 'Mali', 'XOF', '13796354', 'Bamako'),
('MM', 'Myanmar [Burma]', 'MMK', '53414374', 'Naypyitaw'),
('MN', 'Mongolia', 'MNT', '3086918', 'Ulan Bator'),
('MO', 'Macao', 'MOP', '449198', 'Macao'),
('MP', 'Northern Mariana Islands', 'USD', '53883', 'Saipan'),
('MQ', 'Martinique', 'EUR', '432900', 'Fort-de-France'),
('MR', 'Mauritania', 'MRO', '3205060', 'Nouakchott'),
('MS', 'Montserrat', 'XCD', '9341', 'Plymouth'),
('MT', 'Malta', 'EUR', '403000', 'Valletta'),
('MU', 'Mauritius', 'MUR', '1294104', 'Port Louis'),
('MV', 'Maldives', 'MVR', '395650', 'Malé'),
('MW', 'Malawi', 'MWK', '15447500', 'Lilongwe'),
('MX', 'Mexico', 'MXN', '112468855', 'Mexico City'),
('MY', 'Malaysia', 'MYR', '28274729', 'Kuala Lumpur'),
('MZ', 'Mozambique', 'MZN', '22061451', 'Maputo'),
('NA', 'Namibia', 'NAD', '2128471', 'Windhoek'),
('NC', 'New Caledonia', 'XPF', '216494', 'Noumea'),
('NE', 'Niger', 'XOF', '15878271', 'Niamey'),
('NF', 'Norfolk Island', 'AUD', '1828', 'Kingston'),
('NG', 'Nigeria', 'NGN', '154000000', 'Abuja'),
('NI', 'Nicaragua', 'NIO', '5995928', 'Managua'),
('NL', 'Netherlands', 'EUR', '16645000', 'Amsterdam'),
('NO', 'Norway', 'NOK', '5009150', 'Oslo'),
('NP', 'Nepal', 'NPR', '28951852', 'Kathmandu'),
('NR', 'Nauru', 'AUD', '10065', 'Yaren'),
('NU', 'Niue', 'NZD', '2166', 'Alofi'),
('NZ', 'New Zealand', 'NZD', '4252277', 'Wellington'),
('OM', 'Oman', 'OMR', '2967717', 'Muscat'),
('PA', 'Panama', 'PAB', '3410676', 'Panama City'),
('PE', 'Peru', 'PEN', '29907003', 'Lima'),
('PF', 'French Polynesia', 'XPF', '270485', 'Papeete'),
('PG', 'Papua New Guinea', 'PGK', '6064515', 'Port Moresby'),
('PH', 'Philippines', 'PHP', '99900177', 'Manila'),
('PK', 'Pakistan', 'PKR', '184404791', 'Islamabad'),
('PL', 'Poland', 'PLN', '38500000', 'Warsaw'),
('PM', 'Saint Pierre and Miquelon', 'EUR', '7012', 'Saint-Pierre'),
('PN', 'Pitcairn Islands', 'NZD', '46', 'Adamstown'),
('PR', 'Puerto Rico', 'USD', '3916632', 'San Juan'),
('PS', 'Palestine', 'ILS', '3800000', ''),
('PT', 'Portugal', 'EUR', '10676000', 'Lisbon'),
('PW', 'Palau', 'USD', '19907', 'Melekeok'),
('PY', 'Paraguay', 'PYG', '6375830', 'Asunción'),
('QA', 'Qatar', 'QAR', '840926', 'Doha'),
('RE', 'Réunion', 'EUR', '776948', 'Saint-Denis'),
('RO', 'Romania', 'RON', '21959278', 'Bucharest'),
('RS', 'Serbia', 'RSD', '7344847', 'Belgrade'),
('RU', 'Russia', 'RUB', '140702000', 'Moscow'),
('RW', 'Rwanda', 'RWF', '11055976', 'Kigali'),
('SA', 'Saudi Arabia', 'SAR', '25731776', 'Riyadh'),
('SB', 'Solomon Islands', 'SBD', '559198', 'Honiara'),
('SC', 'Seychelles', 'SCR', '88340', 'Victoria'),
('SD', 'Sudan', 'SDG', '35000000', 'Khartoum'),
('SE', 'Sweden', 'SEK', '9828655', 'Stockholm'),
('SG', 'Singapore', 'SGD', '4701069', 'Singapore'),
('SH', 'Saint Helena', 'SHP', '7460', 'Jamestown'),
('SI', 'Slovenia', 'EUR', '2007000', 'Ljubljana'),
('SJ', 'Svalbard and Jan Mayen', 'NOK', '2550', 'Longyearbyen'),
('SK', 'Slovakia', 'EUR', '5455000', 'Bratislava'),
('SL', 'Sierra Leone', 'SLL', '5245695', 'Freetown'),
('SM', 'San Marino', 'EUR', '31477', 'San Marino'),
('SN', 'Senegal', 'XOF', '12323252', 'Dakar'),
('SO', 'Somalia', 'SOS', '10112453', 'Mogadishu'),
('SR', 'Suriname', 'SRD', '492829', 'Paramaribo'),
('SS', 'South Sudan', 'SSP', '8260490', 'Juba'),
('ST', 'São Tomé and Príncipe', 'STD', '175808', 'São Tomé'),
('SV', 'El Salvador', 'USD', '6052064', 'San Salvador'),
('SX', 'Sint Maarten', 'ANG', '37429', 'Philipsburg'),
('SY', 'Syria', 'SYP', '22198110', 'Damascus'),
('SZ', 'Swaziland', 'SZL', '1354051', 'Mbabane'),
('TC', 'Turks and Caicos Islands', 'USD', '20556', 'Cockburn Town'),
('TD', 'Chad', 'XAF', '10543464', 'NDjamena'),
('TF', 'French Southern Territories', 'EUR', '140', 'Port-aux-Français'),
('TG', 'Togo', 'XOF', '6587239', 'Lomé'),
('TH', 'Thailand', 'THB', '67089500', 'Bangkok'),
('TJ', 'Tajikistan', 'TJS', '7487489', 'Dushanbe'),
('TK', 'Tokelau', 'NZD', '1466', ''),
('TL', 'East Timor', 'USD', '1154625', 'Dili'),
('TM', 'Turkmenistan', 'TMT', '4940916', 'Ashgabat'),
('TN', 'Tunisia', 'TND', '10589025', 'Tunis'),
('TO', 'Tonga', 'TOP', '122580', 'Nukualofa'),
('TR', 'Turkey', 'TRY', '77804122', 'Ankara'),
('TT', 'Trinidad and Tobago', 'TTD', '1328019', 'Port of Spain'),
('TV', 'Tuvalu', 'AUD', '10472', 'Funafuti'),
('TW', 'Taiwan', 'TWD', '22894384', 'Taipei'),
('TZ', 'Tanzania', 'TZS', '41892895', 'Dodoma'),
('UA', 'Ukraine', 'UAH', '45415596', 'Kiev'),
('UG', 'Uganda', 'UGX', '33398682', 'Kampala'),
('UM', 'U.S. Minor Outlying Islands', 'USD', '0', ''),
('US', 'United States', 'USD', '310232863', 'Washington'),
('UY', 'Uruguay', 'UYU', '3477000', 'Montevideo'),
('UZ', 'Uzbekistan', 'UZS', '27865738', 'Tashkent'),
('VA', 'Vatican City', 'EUR', '921', 'Vatican City'),
('VC', 'Saint Vincent and the Grenadines', 'XCD', '104217', 'Kingstown'),
('VE', 'Venezuela', 'VEF', '27223228', 'Caracas'),
('VG', 'British Virgin Islands', 'USD', '21730', 'Road Town'),
('VI', 'U.S. Virgin Islands', 'USD', '108708', 'Charlotte Amalie'),
('VN', 'Vietnam', 'VND', '89571130', 'Hanoi'),
('VU', 'Vanuatu', 'VUV', '221552', 'Port Vila'),
('WF', 'Wallis and Futuna', 'XPF', '16025', 'Mata-Utu'),
('WS', 'Samoa', 'WST', '192001', 'Apia'),
('XK', 'Kosovo', 'EUR', '1800000', 'Pristina'),
('YE', 'Yemen', 'YER', '23495361', 'Sanaa'),
('YT', 'Mayotte', 'EUR', '159042', 'Mamoudzou'),
('ZA', 'South Africa', 'ZAR', '49000000', 'Pretoria'),
('ZM', 'Zambia', 'ZMW', '13460305', 'Lusaka'),
('ZW', 'Zimbabwe', 'ZWL', '13061000', 'Harare')
