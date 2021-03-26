if (typeof jQuery !== 'undefined' && typeof saveAs !== 'undefined' && typeof html2canvas !== 'undefined') {
  (function ($) {
    $.fn.wordExport = function (options) {
      var defaultOptions = {
        fileName: 'jQuery-Word-Export', // 文件名称
        hiddenExpr: '[data-print="hide"]', // 需要导出时隐藏的元素表达式
        toImgExpr: '[data-image="true"]', // 需要转成img的元素表达式
        styles: '', // 文档样式
        pageStyles: '{size:595.3pt 841.9pt;margin:36.0pt 36.0pt 36.0pt 36.0pt;mso-header-margin:42.55pt;mso-footer-margin:49.6pt;mso-paper-source:0;}', // 页面样式，页边距之类的
        maxWidth: 624 // 图片最大宽度
      };
      options = $.extend(defaultOptions, options);
      var static = {
        mhtml: {
          top: 'Mime-Version: 1.0\nContent-Base: ' + location.href + '\nContent-Type: Multipart/related; boundary="NEXT.ITEM-BOUNDARY";type="text/html"\n\n--NEXT.ITEM-BOUNDARY\nContent-Type: text/html; charset="utf-8"\nContent-Location: ' + location.href + '\n\n<!DOCTYPE html>\n<html>\n_html_</html>',
          head: '<head>\n<meta http-equiv="Content-Type" content="text/html; charset=utf-8">\n<style>@page WordSection1' + options.pageStyles + '\ndiv.WordSection1{page:WordSection1;}\n_styles_\n</style>\n</head>\n',
          body: '<body><div class=WordSection1>_body_</div></body>'
        }
      };
      // Clone selected element before manipulating it
      var markup = $(this).clone();

      // Remove hidden elements from the output
      markup.each(function () {
        var self = $(this);
        if (self.is(':hidden')) self.remove();
      });

      // Remove hidden elements from the output
      markup.find(':hidden').remove();
      markup.find(options.hiddenExpr).remove();

      var toImg = $(this).find(options.toImgExpr);
      var toImgCopy = markup.find(options.toImgExpr);
      Promise.all(
        toImg.map(function (index, image) {
          return html2canvas(image).then(canvas => {
            let img = document.createElement('IMG');
            img.src = canvas.toDataURL('image/jpeg', 1.0);
            img.dataset.iscanvas = true;
            toImgCopy[index].replaceWith(img);
          });
        })
      ).then(function () {
        // Embed all images using Data URLs
        var images = Array();
        var img = markup.find('img');
        for (var i = 0; i < img.length; i++) {
          // Calculate dimensions of output image
          var w = Math.min(img[i].width, options.maxWidth);
          var h = img[i].height * (w / img[i].width);

          var uri = '';
          if (!img[i].dataset.iscanvas) {
            var canvas = document.createElement('CANVAS');
            canvas.width = w;
            canvas.height = h;
            // Draw image to canvas
            var context = canvas.getContext('2d');
            context.drawImage(img[i], 0, 0, w, h);
            // Get data URL encoding of image
            uri = canvas.toDataURL('image/png');
            $(img[i]).attr('src', uri);
          } else {
            uri = img[i].src;
          }

          img[i].width = w;
          img[i].height = h;
          // Save encoded image to array
          images[i] = {
            type: uri.substring(uri.indexOf(':') + 1, uri.indexOf(';')),
            encoding: uri.substring(uri.indexOf(';') + 1, uri.indexOf(',')),
            location: $(img[i]).attr('src'),
            data: uri.substring(uri.indexOf(',') + 1)
          };
        }

        // Prepare bottom of mhtml file with image data
        var mhtmlBottom = '\n';
        for (var i = 0; i < images.length; i++) {
          mhtmlBottom += '--NEXT.ITEM-BOUNDARY\n';
          mhtmlBottom += 'Content-Location: ' + images[i].location + '\n';
          mhtmlBottom += 'Content-Type: ' + images[i].type + '\n';
          mhtmlBottom += 'Content-Transfer-Encoding: ' + images[i].encoding + '\n\n';
          mhtmlBottom += images[i].data + '\n\n';
        }
        mhtmlBottom += '--NEXT.ITEM-BOUNDARY--';

        // Aggregate parts of the file together
        var fileContent = static.mhtml.top.replace('_html_', static.mhtml.head.replace('_styles_', options.styles) + static.mhtml.body.replace('_body_', markup.html())) + mhtmlBottom;

        // Create a Blob with the file contents
        var blob = new Blob([fileContent], {
          type: 'application/msword;charset=utf-8'
        });
        saveAs(blob, options.fileName + '.doc');
      });
    };
  })(jQuery);
} else {
  if (typeof jQuery === 'undefined') {
    console.error('jQuery Word Export: missing dependency (jQuery)');
  }
  if (typeof saveAs === 'undefined') {
    console.error('jQuery Word Export: missing dependency (FileSaver.js)');
  }
  if (typeof html2canvas === 'undefined') {
    console.error('jQuery Word Export: missing dependency (html2canvas.js)');
  }
}
