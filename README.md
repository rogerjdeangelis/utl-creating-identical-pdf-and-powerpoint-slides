# utl-creating-identical-pdf-and-powerpoint-slides
Creating identical pdf and powerpoint_slides.
    Creating identical pdf and powerpoint_slides

         Three Tasks

            1. Pdf slides
            2. Matching Powerpoint slides
            3. Adding a theme to powerpoint slides (google'powerpoint photo album transparancy'
               I am not a big fan of mouse surfing.

    pdf output
    https://tinyurl.com/yd9pczfv
    https://github.com/rogerjdeangelis/utl-creating-identical-pdf-and-powerpoint-slides/blob/master/utl-creating-identical-pdf-powerpoint_slides.pdf

    powerpoint output
    https://tinyurl.com/y9pjwjh2
    https://github.com/rogerjdeangelis/utl-creating-identical-pdf-and-powerpoint-slides/blob/master/utl-creating-identical-pdf-powerpoint_slides.pptx

    github
    https://github.com/rogerjdeangelis/utl-creating-identical-pdf-and-powerpoint-slides

    INPUT
    =====

     You need to create directories

     d:/pdf/inp

     and download (or install ghostscript http://www.ghostscript.com/download/)

     I copied the executable gswin64c.exe to my d:\pdf\inp folder
     C:\Program Files\gs\gs9.25\bin\gswin64c.exe to d:\pdf\gswin64c.exe
     There is a 32bit version.

       d:/pdf/inp

           gswin64c.exe

     SASHELP.CLASS total obs=19

      NAME       SEX    AGE    HEIGHT    WEIGHT

      Alfred      M      14     69.0      112.5
      Alice       F      13     56.5       84.0
      Barbara     F      13     65.3       98.0
      Carol       F      14     62.8      102.5
      Henry       M      14     63.5      102.5
      James       M      12     57.3       83.0
     ....


    EXAMPLE OUTPUT (exactly the same in PDF and Powerpoint)
    =======================================================

     d:/pdf/utl-creating-identical-pdf-powerpoint_slides.pdf


     Links       +----------------------------------------------------------------+
                 |                                                                |
                 |                                                                |
     + Males     |          __  __       _   _        ____ _                      |
                 |         |  \/  | __ _| |_| |__    / ___| | __ _ ___ ___        |
        - Age    |         | |\/| |/ _` | __| '_ \  | |   | |/ _` / __/ __|       |
                 |         | |  | | (_| | |_| | | | | |___| | (_| \__ \__ \       |
                 |         |_|  |_|\__,_|\__|_| |_|  \____|_|\__,_|___/___/       |
     + Females   |                                                                |
                 |                                                                |
        - Age    |          by Roger J Deangelis                                  |
                 |                                                                |
                 |                                                                |
                 |                                                                |
                 |                                                                |
                 +----------------------------------------------------------------+


                 +----------------------------------------------------------------+
                 |                                                                |
                 |       ____            _             _                          |
                 |      / ___|___  _ __ | |_ ___ _ __ | |_ ___                    |
                 |     | |   / _ \| '_ \| __/ _ \ '_ \| __/ __|                   |
                 |     | |__| (_) | | | | ||  __/ | | | |_\__ \                   |
                 |      \____\___/|_| |_|\__\___|_| |_|\__|___/                   |
                 |                                                                |
                 |                                                                |
                 |      Males .......................................... 1        |
                 |                                                                |
                 |        Age .......................................... 1        |
                 |                                                                |
                 |                                                                |
                 |      Females ........................................ 2        |
                 |                                                                |
                 |        Age .......................................... 2        |
                 |                                                                |
                 |                                                                |
                 +----------------------------------------------------------------+


                 +----------------------------------------------------------------+
                 |                                                                |
                 |                __  __       _                                  |
                 |               |  \/  | __ _| | ___  ___                        |
                 |               | |\/| |/ _` | |/ _ \/ __|                       |
                 |               | |  | | (_| | |  __/\__ \                       |
                 |               |_|  |_|\__,_|_|\___||___/                       |
                 |                                                                |
                 |                                                                |
                 |                                                                |
                 |      +-------------------------------------------------+       |
                 |      | NAME     |   SEX  |    AGE  |  HEIGHT |  WEIGHT |       |
                 |      +----------+--------+---------+---------+---------+       |
                 |      | ALFRED   |    M   |    14   |    69   |  112.5  |       |
                 |      +----------+--------+---------+---------+---------+       |
                 |       ...                                                      |
                 |      +----------+--------+---------+---------+---------+       |
                 |      | WILLIAM  |    M   |    15   |   66.5  |  112    |       |
                 |      +----------+--------+---------+---------+---------+       |
                 |                                                                |
                 |                                                                |
                 |                                                                |
                 +----------------------------------------------------------------+


                 +----------------------------------------------------------------+
                 |                                                                |
                 |                                                                |
                 |         *_____                    _                            |
                 |         |  ___|__ _ __ ___   __ _| | ___  ___                  |
                 |         | |_ / _ \ '_ ` _ \ / _` | |/ _ \/ __|                 |
                 |         |  _|  __/ | | | | | (_| | |  __/\__ \                 |
                 |         |_|  \___|_| |_| |_|\__,_|_|\___||___/                 |
                 |                                                                |
                 |                                                                |
                 |      +-------------------------------------------------+       |
                 |      | NAME     |   SEX  |    AGE  |  HEIGHT |  WEIGHT |       |
                 |      +----------+--------+---------+---------+---------+       |
                 |      | ALICE    |    F   |    15   |    59   |  102.5  |       |
                 |      +----------+--------+---------+---------+---------+       |
                 |       ...                                                      |
                 |      +----------+--------+---------+---------+---------+       |
                 |      | BARBARA  |    F   |    15   |   63.5  |   92    |       |
                 |      +----------+--------+---------+---------+---------+       |
                 |                                                                |
                 |                                                                |
                 |                                                                |
                 +----------------------------------------------------------------+



    PROCESS
    =======

    1. Pdf slides
    -------------

     title;
     footnote;

     %utl_pptlan100(topmargin=1in);

     options label orientation=landscape;

     ods pdf contents style=utl_pptlan100
         file="d:/pdf/inp/utl-creating-identical-pdf-powerpoint_slides.pdf";

     ods proclabel="Males";

     proc report data=sashelp.class(where=(sex="M")) contents="Age"
        style(report)={outputwidth=100pct font_size=15pt};
     cols sex age height weight;
     define sex / group;
     define age / display;
     break before sex / contents="" page;
     run;quit;

     ods proclabel="Females";
     proc report data=sashelp.class(where=(sex="F")) contents="Age"
        style(report)={outputwidth=100pct font_size=15pt};
     cols  sex age height weight;
     define sex / group;
     define age / display;
     break before sex / contents="" page;
     run;quit;

     ods pdf close;


    2. Powerpoint slides
    --------------------

    * pages to jpegs;
    x "cd d:/pdf/inp";
    x "gswin64c.exe -dNOPAUSE -dBATCH -dSAFER -dGraphicsAlphaBits=4 -dTextAlphaBits=4 -sDEVICE=jpeg -r300
       -sOutputFile=page-%00d.jpg utl-creating-identical-pdf-powerpoint_slides.pdf";
    x "cd c:/utl";

    * creates jpegs

      d:/pdf/inp

       page-1.pdf
       page-2.pdf
       page-3.pdf



    /*
    * you can also create BMPs.
    x "cd d:/pdf/inp";
    x "gswin64c.exe -dNOPAUSE -dBATCH -dSAFER -sDEVICE=bmp256 -dTextAlphaBits=4 -sOutputFile=page-%00d.bmp three_pages.pdf";
    x "cd c:/utl";
    */

    * insert jpegs;
    Open powerpoint and insert jpeg slides

      insert-> photo album _> new photo album -> select directory -> shift-select-first-picture next shift-select-last-picture -> create


    OUTPUT
    ======

    see above


    *                _               _       _
     _ __ ___   __ _| | _____     __| | __ _| |_ __ _
    | '_ ` _ \ / _` | |/ / _ \   / _` |/ _` | __/ _` |
    | | | | | | (_| |   <  __/  | (_| | (_| | || (_| |
    |_| |_| |_|\__,_|_|\_\___|   \__,_|\__,_|\__\__,_|

    ;

     You need to create directories

     d:/pdf/inp

     and download (or install ghostscript http://www.ghostscript.com/download/)

     I copied the executable gswin64c.exe to my d:\pdf\inp folder
     C:\Program Files\gs\gs9.25\bin\gswin64c.exe to d:\pdf\gswin64c.exe
     There is a 32bit version.

       d:/pdf/inp

           gswin64c.exe

     SASHELP.CLASS total obs=19
