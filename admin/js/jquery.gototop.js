(function($) {
  $.fn.gototop = function(opts) {
	var defaults = {
	  startFadeIn: 100,
	  container: "#container",
	  toTop: "#toTop"
	}
    var options = $.extend(defaults, opts);
	var items;
	if (this.length === 0) {
        $(options.container).prepend('<div id="toTop">toTop</div>');
		items = $(options.toTop); 
	}
	else {
		items = this;
	}
    // Plugin code
    return items.each(function() {
        var $top = $(this);
        function fnScroll() {
		  if ($(window).scrollTop() > options.startFadeIn) {
			$top.fadeIn().css("display", "inline-block");	
		  } else {
			$top.fadeOut();
		  }
		};          
        $(window).load(fnScroll).scroll(fnScroll);
	    $top.click(function() {
		    $('body,html').animate({scrollTop:0},800);
		});
	});
  }
})(jQuery);
