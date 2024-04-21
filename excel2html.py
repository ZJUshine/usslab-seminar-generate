# 遍历Excel中的数据，生成HTML模板
import pandas as pd
import gdown

# 下载Excel文件
excel_path = 'https://drive.google.com/uc?id=1mpHEXUUKO-8wjuAiwNW3Y350jmzO41x-'
gdown.download(excel_path, "2024_Spring_Summer.xlsx", quiet=False)

# 读取Excel文件
excel_file = "./2024_Spring_Summer.xlsx"
df = pd.read_excel(excel_file, header=None, skiprows=1,engine='openpyxl')
df.columns = ['日期', '姓名', '文章', '会议']
df.drop(0,inplace=True) # 删除相同的'日期', '姓名', '文章', '会议'行

html_template = "" # 初始化模板
df['日期'].fillna(method='ffill', inplace=True)
df.fillna('', inplace=True)
# 根据连续的相同的“日期”进行分组
groups = (df['日期'] != df['日期'].shift()).cumsum()
# df['group_index'] = df.groupby('日期').cumcount()
# df.reset_index(drop=True, inplace=True)
df_grouped = df.groupby(groups)


html_template += '''
<!DOCTYPE html>
<html>

<head>
    <!-- Standard Meta -->
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
    <base href="/" />
    <!-- Site Properties -->
    <title>USS Lab. - Ubiquitous System Security Lab.</title>
    <link rel="stylesheet" type="text/css" href="semantic/dist/semantic.min.css" />
    <link rel="shortcut icon" type="image/x-icon" href="images/favicon.ico" />
    <link rel="stylesheet" href="third-party/Swiper-3.4.1/dist/css/swiper.min.css" />

    <style type="text/css">
        .hidden.menu {
            display: none;
        }

        .largeheader {
            margin-top: 3em;
            margin-bottom: 0em;
            font-size: 4em;
            font-weight: normal;
        }

        .ui.vertical.stripe {
            padding: 3em 0em;
        }

        .ui.vertical.stripe h3 {
            font-size: 2em;
        }

        .ui.vertical.stripe .button+h3,
        .ui.vertical.stripe p+h3 {
            margin-top: 3em;
        }

        .ui.vertical.stripe .floated.image {
            clear: both;
        }

        .ui.vertical.stripe p {
            font-size: 1.33em;
        }

        .ui.vertical.stripe .horizontal.divider {
            margin: 3em 0em;
        }

        .quote.stripe.segment {
            padding: 0em;
        }

        .quote.stripe.segment .grid .column {
            padding-top: 5em;
            padding-bottom: 5em;
        }

        .footer.segment {
            padding: 5em 0em;
        }

        .secondary.pointing.menu .toc.item {
            display: none;
        }

        .top.fixed.menu .ftoc.item {
            display: none;
        }

        @media only screen and (max-width: 700px) {
            .ui.fixed.menu {
                display: none !important;
            }

            .secondary.pointing.menu .item,
            .secondary.pointing.menu .logo,
            .secondary.pointing.menu .menu {
                display: none;
            }

            .secondary.pointing.menu .toc.item {
                display: block;
            }

            .masthead.segment {
                min-height: 350px;
            }

            .masthead h1.ui.header {
                font-size: 2em;
                margin-top: 1.5em;
            }

            .masthead h2 {
                margin-top: 0.5em;
                font-size: 1.5em;
            }
        }

        @media only screen and (max-width: 700px) {
            .ui.fixed.menu {
                display: none !important;
            }

            .top.fixed.menu .item,
            .top.fixed.menu .menu {
                display: none;
            }

            .top.fixed.menu .ftoc.item {
                display: block;
            }
        }

        .logo.image {
            width: 50px;
            margin-right: 1px;
        }

        .galary {
            margin-top: 20px;
        }

        .swiper-container {
            width: 100%;
            height: 600px;
        }

        .beian:hover {
            color: white;
        }

        .beian {
            color: grey;
        }
    </style>

    <script src="//cdn.bootcss.com/jquery/3.1.1/jquery.min.js"></script>
    <script src="semantic/dist/components/visibility.js"></script>
    <script src="semantic/dist/components/sidebar.js"></script>
    <script src="semantic/dist/components/transition.js"></script>
    <script src="third-party/Swiper-3.4.1/dist/js/swiper.min.js"></script>
    <script>
        $(document)
            .ready(function () {

                // fix menu when passed
                $('.secondary.menu')
                    .visibility({
                        once: false,
                        onBottomPassed: function () {
                            $('.fixed.menu').transition('fade in');
                        },
                        onBottomPassedReverse: function () {
                            $('.fixed.menu').transition('fade out');
                        }
                    });

                // create sidebar and attach to menu open
                $('.ui.sidebar').sidebar('attach events', '.toc.item');
                $('.ui.sidebar').sidebar('attach events', '.ftoc.item');

                var swiper = new Swiper('.swiper-container', {
                    pagination: '.swiper-pagination',
                    paginationClickable: true,
                    speed: 1000,
                    centeredSlides: true,
                    autoplay: 2500,
                });

            });
    </script>
</head>

<body>

    <!-- Following Menu -->
    <div class="ui large top fixed hidden menu">
        <div class="ui container">
            <a class="ftoc item">
                <i class="sidebar icon"></i>
            </a>
            <a class="header item">
                <img class="logo" src="images/logo.png" /> USSLab.
            </a>
            <a class="item" href="/">Home</a>
            <a class="item" href="projects/">Projects</a>
            <a class="item" href="publications.html">Publications</a>
            <a class="item" href="member/">Members</a>
            <a class="item active">Courses</a>
            <a class="item" href="/jarvisoj.html">Jarvis OJ</a>
            <a class="item active" href="/seminar.html">Seminar</a>

            <a class="right item" href="/contact.html">Contact us</a>


        </div>
    </div>

    <!-- Sidebar Menu -->
    <div class="ui vertical inverted sidebar menu">
        <div class="item">
            <a class="ui logo icon image">
                <img src="images/logo.png" />
            </a>
            <a href="#"><b>USSLab.</b></a>
        </div>
        <a class="item" href="/">Home</a>
        <a class="item" href="projects/">Projects</a>
        <a class="item" href="publications.html">Publications</a>
        <a class="item" href="member/">Members</a>
        <a class="item active">Courses</a>
        <a class="item" href="/jarvisoj.html">Jarvis OJ</a>
        <a class="item active" href="/seminar.html">Seminar</a>

        <a class="right item" href="/contact.html">Contact us</a>
    </div>


    <!-- Page Contents -->
    <div class="pusher">
        <div class="ui inverted vertical center aligned segment">
            <div class="ui container">
                <div class="ui large secondary inverted pointing menu">
                    <a class="toc item">
                        <i class="sidebar icon"></i>
                    </a>
                    <a class="ui logo icon image">
                        <img src="images/logo.png" />
                    </a>
                    <a class="item" href="/"><b>USS Lab.</b></a>
                    <a class="item" href="/">Home</a>
                    <a class="item" href="projects/">Projects</a>
                    <a class="item" href="publications.html">Publications</a>
                    <a class="item" href="member/">Members</a>
                    <a class="item" href="/courses.html">Courses</a>
                    <a class="item" href="/jarvisoj.html">Jarvis OJ</a>
                    <a class="item active" href="/seminar.html">Seminar</a>
                    <a class="right item" href="/contact.html">Contact us</a>
                </div>
            </div>

        </div>

        <div class="ui vertical stripe container segment">
            <h1 class="ui header">Seminar</h1>


            <div class="ui divider"></div>



            <h2 class="ui header">Spring 2023 Seminars:</h2>
            <div>
                <table class="table tiny-table" rules=rows width="100%">
                    <tr style="color:#FFFFFF;text-decoration:none;background:#005A9C;text-align:center;font-weight:bold;"
                        valign="top">
                        <td width="15%"></td>
                        <td align="left" width="15%">
                            <p>Speaker</p>
                        </td>
                        <td align="left" width="50%">
                            <p>Title</p>
                        </td>
                        <td align="left" width="20%">
                            <p>Conference</p>

                        </td>
                    </tr>

                    <tr style="color:#FFFFFF;text-decoration:none;background:#005A9C;text-align:center;font-weight:bold;"
                        valign="top">
                        <td width="15%"></td>
                        <td align="left" width="15%"></td>
                        <td align="left" width="50%"></td>
                        <td align="left" width="20%"></td>
                    </tr>'''
html_template += '''

'''

for num, group in df_grouped:
    group_num = len(group)
    group.reset_index(drop=True, inplace=True)
    for idx, row in group.iterrows():
        # 这里写表头
        if idx == 0:
            formatted_date = row['日期'].strftime('%Y/%m/%d') # 格式化为 'YYYY/MM/DD' 格式

            html_template += f'''
                    <tr style="font-size:15px" valign="center">
                        <td align="center" rowspan="{group_num}">{formatted_date}</td>
                '''
        else:
            html_template += '''
                    <tr style="font-size:15px" valign="center">
            '''
        
        # 开始写姓名，文章
        html_template += f'''
                        <td align="left" width="15%" style="font-weight:bold;">{row['姓名']}</td>
                        <td align="left" width="50%">{row['文章']}</td>
                        <td align="left" width="20%">{row['会议']}</td>
                    </tr>
        '''

    # group之间的换行
    html_template += f'''

    '''
    # html_template += '''<tr style="font-size:15px" valign="center">'''
    # html_template += '''<tr style="font-size:15px" valign="center">'''

    # <tr style="font-size:15px" valign="center">
    #                     <td align="center" rowspan="3">2023/9/15</td>
    #                     <td align="left" width="15%" style="font-weight:bold;">Zhouhao Ji</td>
    #                     <td align="left" width="50%">Talk: MaDIoT Attack</td>
    #                     <td align="left" width="20%"></td>
    #                 </tr>
    #                 <tr style="font-size:15px" valign="center">
    #                     <td align="left" width="15%" style="font-weight:bold;">Zhicong Zheng</td>
    #                     <td align="left" width="50%">Fuzzing Hardware Like Software</td>
    #                     <td align="left" width="20%">USENIX 2022</td>
    #                 </tr>
    #                 <tr style="font-size:15px" valign="center">
    #                     <td align="left" width="15%" style="font-weight:bold;">Namin Hou</td>
    #                     <td align="left" width="50%">MagBackdoor: Beware of Your Loudspeaker as A Backdoor For Magnetic Injection Attacks</td>
    #                     <td align="left" width="20%">S&P 2023</td>
                    # </tr>

html_template += f'''                </table>
                <hr>
            </div>

            <!-- -------------split line------------------- -->
            <h2 class="ui header"><a href="seminars/seminar-2023-autumn.html">Autumn 2023 Seminars</a></h2>
            <h2 class="ui header"><a href="seminars/seminar-2023-spring.html">Spring 2023 Seminars</a></h2>
            <h2 class="ui header"><a href="seminars/seminar-2022-autumn.html">Autumn 2022 Seminars</a></h2>
            <h2 class="ui header"><a href="seminars/seminar-2022-spring.html">Spring 2022 Seminars</a></h2>
            <h2 class="ui header"><a href="seminars/seminar-2021-autumn.html">Autumn 2021 Seminars</a></h2>
            <h2 class="ui header"><a href="seminars/seminar-2021-spring.html">Spring 2021 Seminars</a></h2>
            <h2 class="ui header"><a href="seminars/seminar-2020-autumn.html">Autumn 2020 Seminars</a></h2>
            <h2 class="ui header"><a href="seminars/seminar-2020-summer.html">Summer 2020 Seminars</a></h2>
            <h2 class="ui header"><a href="seminars/seminar-2020-spring.html">Spring 2020 Seminars</a></h2>
            <h2 class="ui header"><a href="seminars/seminar-2019-autumn.html">Autumn 2019 Seminars</a></h2>
            <h2 class="ui header"><a href="seminars/seminar-2019-spring.html">Spring 2019 Seminars</a></h2>
            <h2 class="ui header"><a href="seminars/seminar-2018-autumn.html">Autumn 2018 Seminars</a></h2>
        </div>
            

        <div class="ui inverted vertical footer segment">
            <div class="ui center aligned container">
                <!--<img src="assets/images/logo.png" class="ui centered mini image">-->
                <div class="ui horizontal inverted small divided link list">
                    Copyright &copy; USS Lab. 2016-2017
                    <br /> Zhejiang University, Hangzhou, Zhejiang, China. 310027
                    <br /> All Right Reserved.
                    <br />
                    <a class="beian" href="http://www.miitbeian.gov.cn/" target="_blank">
                        <b>浙ICP备16045554号-1</b>
                    </a>
                </div>
            </div>
        </div>
    </div>

</body>

</html>
'''

# 将生成的HTML保存到文件或进行其他操作
with open("seminar.html", "w", encoding='utf-8') as f:
    f.write(html_template)

print("HTML模板生成完成，并保存到seminar.html")
