if (typeof jQuery !== "undefined" && typeof saveAs !== "undefined") {
    (function($) {
        $.fn.wordExport = function(fileName) {
            fileName = typeof fileName !== 'undefined' ? fileName : "jQuery-Word-Export";
            var static = {
                mhtml: {
                    top: "Mime-Version: 1.0\nContent-Base: " + location.href + "\nContent-Type: Multipart/related; boundary=\"NEXT.ITEM-BOUNDARY\";type=\"text/html\"\n\n--NEXT.ITEM-BOUNDARY\nContent-Type: text/html; charset=\"utf-8\"\nContent-Location: " + location.href + "\n\n<!DOCTYPE html>\n<html>\n_html_</html>",
                    head: "<head>\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\n<style>\n_styles_\n</style>\n</head>\n",
                    body: "<body>_body_</body>"
                }
            };
            var options = {
                maxWidth: 624
            };
            // Clone selected element before manipulating it
            var markup = $(this).clone();

            // Remove hidden elements from the output
            markup.each(function() {
                var self = $(this);
                if (self.is(':hidden'))
                    self.remove();
            });

            // Embed all images using Data URLs
            var images = Array();
            var img = markup.find('img');
            for (var i = 0; i < img.length; i++) {
                // Calculate dimensions of output image
                var w = Math.min(img[i].width, options.maxWidth);
                var h = img[i].height * (w / img[i].width);
                // Create canvas for converting image to data URL
                $('<canvas>').attr("id", "jQuery-Word-export_img_" + i).width(w).height(h).insertAfter(img[i]);
                var canvas = document.getElementById("jQuery-Word-export_img_" + i);
                canvas.width = w;
                canvas.height = h;
                // Draw image to canvas
                var context = canvas.getContext('2d');
                context.drawImage(img[i], 0, 0, w, h);
                // Get data URL encoding of image
                var uri = canvas.toDataURL();
                $(img[i]).attr("src", img[i].src);
                img[i].width = w;
                img[i].height = h;
                // Save encoded image to array
                images[i] = {
                    type: uri.substring(uri.indexOf(":") + 1, uri.indexOf(";")),
                    encoding: uri.substring(uri.indexOf(";") + 1, uri.indexOf(",")),
                    location: $(img[i]).attr("src"),
                    data: uri.substring(uri.indexOf(",") + 1)
                };
                // Remove canvas now that we no longer need it
                canvas.parentNode.removeChild(canvas);
            }

            // Prepare bottom of mhtml file with image data
            var mhtmlBottom = "\n";
            for (var i = 0; i < images.length; i++) {
                mhtmlBottom += "--NEXT.ITEM-BOUNDARY\n";
                mhtmlBottom += "Content-Location: " + images[i].contentLocation + "\n";
                mhtmlBottom += "Content-Type: " + images[i].contentType + "\n";
                mhtmlBottom += "Content-Transfer-Encoding: " + images[i].contentEncoding + "\n\n";
                mhtmlBottom += images[i].contentData + "\n\n";
            }
            mhtmlBottom += "--NEXT.ITEM-BOUNDARY--";

            //TODO: load css from included stylesheet
            var data = [[1,"Phạm Lâm",40,7,3,2],[2,"Minh trọc",15,3,3,0]] ;
            var styles = "";
            var a = "hello";
            var b = "" ;
            var stylesheet2 = "style='border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .75pt;padding:0in 0in 0in 0in'";
            for(i = 0;i< data.length; i++){
                b += "<tr style='mso-yfti-irow:3'><td " + stylesheet2 + "><p class=MsoNormal align=center style='margin-top:1.5pt;margin-right:1.5pt;margin-bottom:1.5pt;margin-left:2.0pt;text-align:center'>" + data[i][0] + "</p></td><td " + stylesheet2 + "<p class=MsoNormal align=center style='margin-top:1.5pt;margin-right:1.5pt;margin-bottom:1.5pt;margin-left:2.0pt;text-align:center'>" + data[i][1] + "</p></td><td " + stylesheet2 + "<p class=MsoNormal align=center style='mso-margin-top-alt:auto;margin-right:1.5pt;mso-margin-bottom-alt:auto;margin-left:2.0pt;text-align:center'>" + data[i][2] + "</p></td><td " + stylesheet2 + "<p class=MsoNormal align=center style='margin-top:1.5pt;margin-right:1.5pt;margin-bottom:1.5pt;margin-left:2.0pt;text-align:center'>" + data[i][3] + "</p></td><td " + stylesheet2 + "<p class=MsoNormal align=center style='margin-top:1.5pt;margin-right:1.5pt;margin-bottom:1.5pt;margin-left:2.0pt;text-align:center'>" + data[i][4] + "</p></td><td " + stylesheet2 + "<p class=MsoNormal align=center style='margin-top:1.5pt;margin-right:1.5pt;margin-bottom:1.5pt;margin-left:2.0pt;text-align:center'>" + data[i][5] + "</p></td></tr>";
            }
            var stylesheet = "border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .75pt;padding:0in 0in 0in 0in";


            var booksToRead = "<table border=0 cellpadding=0 style='width: 100%;mso-cellspacing:1.5pt'>" +
                "<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'><p class=MsoNormal align=center style='text-align:center'><b>UBND TỈNH QUẢNG NINH<o:p></o:p></b></p></td><td style='padding:.75pt .75pt .75pt .75pt'><p class=MsoNormal align=center style='text-align:center'><b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM<o:p></o:p></b></p> </td> </tr><tr style='mso-yfti-irow:1'><td style='padding:.75pt .75pt .75pt .75pt'><p class=MsoNormal align=center style='text-align:center'><b>VĂN PHÒNG<o:p></o:p></b></p></td><td style='padding:.75pt .75pt .75pt .75pt'><p class=MsoNormal align=center style='text-align:center'><b>Độc lập - Tự do - Hạnh phúc<o:p></o:p></b></p></td></tr><tr style='mso-yfti-irow:2;height:.75pt'><td style='padding:.75pt .75pt .75pt .75pt;height:.75pt'><p class=MsoNormal align=center style='text-align:center;mso-line-height-alt:.75pt'>___________</p></td><td style='padding:.75pt .75pt .75pt .75pt;height:.75pt'><p class=MsoNormal align=center style='text-align:center;mso-line-height-alt:.75pt'>_______________________</p></td></tr><tr style='mso-yfti-irow:3;mso-yfti-lastrow:yes'><td style='padding:.75pt .75pt .75pt .75pt'><p class=MsoNormal><o:p>&nbsp;</o:p></p></td><td style='padding:.75pt .75pt .75pt .75pt'><p class=MsoNormal align=right style='text-align:right'><i>Quảng Ninh, ngày 17 tháng 09 năm 2014</i></p></td></tr>" +

                "</table>" +
                "<table class=MsoNormalTable border=0 cellpadding=0 style='width: 100%;mso-cellspacing:1.5pt' " +
                "<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'><td style='padding:.75pt .75pt .75pt .75pt'><p class=MsoNormal>&nbsp;</p></td></tr><tr style='mso-yfti-irow:1'><td style='padding:.75pt .75pt .75pt .75pt'><p class=MsoNormal align=center style='text-align:center'><b>KẾT QUẢ XỬ LÝ CÔNG VIỆC <o:p></o:p></b></p></td></tr><tr style='mso-yfti-irow:2'><td style='padding:.75pt .75pt .75pt .75pt'><p class=MsoNormal align=center style='text-align:center'><b>(Tên thông tin cần tìm:…….)<o:p></o:p></b></p></td></tr><tr style='mso-yfti-irow:3;mso-yfti-lastrow:yes'><td style='padding:.75pt .75pt .75pt .75pt'><p class=MsoNormal align=center style='text-align:center'><o:p>&nbsp;</o:p></p></td></tr>" +
                "</table>" +
                "<table class=MsoNormalTable border=1 cellspacing=0 cellpadding=0 style='width:90%;mso-cellspacing:0in;border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .75pt;mso-padding-alt:0in 0in 0in 0in;mso-border-insideh:.75pt solid windowtext;mso-border-insidev:.75pt solid windowtext'>" +
                "<thead><tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'><td rowspan='2' style='width: 5%; " + stylesheet + "'<p class=MsoNormal align=center style='text-align:center'><b>TT<o:p></o:p></b></p></td><td rowspan='2' style='width: 29%;" + stylesheet + "'><p class=MsoNormal align=center style='text-align:center'><b>Họ và tên<o:p></o:p></b></p></td><td rowspan='2' style='width: 11%; " + stylesheet + "'><p class=MsoNormal align=center style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;text-align:center'><b>Tổng số trang<o:p></o:p></b></p></td><td rowspan='2' style='width: 11%;" + stylesheet + "'><p class=MsoNormal align=center style='text-align:center'><b>Tổng số file<o:p></o:p></b></p></td><td colspan='2' style='width: 44%;" + stylesheet + "'><p class=MsoNormal align=center style='text-align:center'><b>Tình trạng xử lý<o:p></o:p></b></p></td></tr><tr style='mso-yfti-irow:1'><td style='width:19.32%;border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .75pt;padding:0in 0in 0in 0in'><p align=center style='text-align:center'><b>Đã cập nhật đầy đủ thông tin<o:p></o:p></b></p></td><td style='width:24.7%;border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .75pt;padding:0in 0in 0in 0in'><p align=center style='text-align:center'><b>Cập nhật ko đầy đủ thông tin<o:p></o:p></b></p></td></tr></thead>" +
                "<tbody>" +
                b +
                "<tr style='mso-yfti-irow:4;mso-yfti-lastrow:yes'><td colspan=2 style='width:33.3%;border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .75pt;padding:0in 0in 0in 0in'><p class=MsoNormal align=center style='mso-margin-top-alt:auto;margin-right:1.5pt;mso-margin-bottom-alt:auto;margin-left:2.0pt;text-align:center'>TỔNG CỘNG</p></td><td style='width:11.66%;border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .75pt;padding:0in 0in 0in 0in'><p class=MsoNormal align=center style='mso-margin-top-alt:auto;margin-right:1.5pt;mso-margin-bottom-alt:auto;margin-left:2.0pt;text-align:center'>45</p></td><td valign=top style='width:11.02%;border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .75pt;padding:0in 0in 0in 0in'><p class=MsoNormal align=center style='mso-margin-top-alt:auto;margin-right:1.5pt;mso-margin-bottom-alt:auto;margin-left:2.0pt;text-align:center'>10</p></td><td style='width:19.32%;border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .75pt;padding:0in 0in 0in 0in'><p class=MsoNormal align=center style='mso-margin-top-alt:auto;margin-right:1.5pt;mso-margin-bottom-alt:auto;margin-left:2.0pt;text-align:center'>6</p></td><td style='width:24.7%;border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .75pt;padding:0in 0in 0in 0in'><p class=MsoNormal align=center style='mso-margin-top-alt:auto;margin-right:1.5pt;mso-margin-bottom-alt:auto;margin-left:2.0pt;text-align:center'>2</p></td></tr></tr>" +
                "</tbody>" +
                "</table>" +
                "<p class=MsoNormal style='margin-bottom:12.0pt'><o:p>&nbsp;</o:p></p>" +
                "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='width:100.0%;mso-cellspacing:0in;mso-padding-alt:0in 0in 0in 0in'>" +
                "<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'><td style='width:70.0%;padding:0in 0in 0in 0in'><p class=MsoNormal>&nbsp;</p></td><td style='width:25.0%;padding:0in 0in 0in 0in'><p class=MsoNormal align=center style='text-align:center'><b>NGƯỜI TỔNG HỢP<o:p></o:p></b></p></td><td style='padding:0in 0in 0in 0in'><p class=MsoNormal>&nbsp;</p></td></tr><tr style='mso-yfti-irow:1'><td style='padding:0in 0in 0in 0in'><p class=MsoNormal>&nbsp;</p></td><td style='padding:0in 0in 0in 0in'><p class=MsoNormal align=center style='text-align:center'>&nbsp;</p></td><td style='padding:0in 0in 0in 0in'><p class=MsoNormal>&nbsp;</p></td></tr><tr style='mso-yfti-irow:2;mso-yfti-lastrow:yes'><td style='padding:0in 0in 0in 0in'><p class=MsoNormal>&nbsp;</p></td><td style='padding:0in 0in 0in 0in'><p class=MsoNormal align=center style='text-align:center'>Nguyễn Tấn Dũng</p></td><td style='padding:0in 0in 0in 0in'><p class=MsoNormal>&nbsp;</p></td></tr>" +
                "</table>";

            // Aggregate parts of the file together 
            var fileContent = static.mhtml.top.replace("_html_",static.mhtml.head.replace("_styles_", styles) + static.mhtml.body.replace("_body_",booksToRead)) + mhtmlBottom;

            // Create a Blob with the file contents
            var blob = new Blob([fileContent], {
                type: "application/msword;charset=utf-8"
            });
            saveAs(blob, fileName + ".doc");
        };
    })(jQuery);
} else {
    if (typeof jQuery === "undefined") {
        console.error("jQuery Word Export: missing dependency (jQuery)");
    }
    if (typeof saveAs === "undefined") {
        console.error("jQuery Word Export: missing dependency (FileSaver.js)");
    };
}

// "<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'><th style='width: 5%; " + stylesheet+"'></th><th style='width: 25%; " + stylesheet + "'></th><th colspan='3' style='width: 40%;'><b>Tình trạng xử lý</b></th></tr>" +
