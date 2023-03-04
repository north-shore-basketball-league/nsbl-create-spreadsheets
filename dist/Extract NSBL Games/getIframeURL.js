iframes = document.querySelectorAll('[title="Table Master"]');

output = [];

iframes.forEach((i) => {
	if (i.getAttribute("src") === null) return null;

	output.push('"' + i.getAttribute("src") + '"');
});

return "[" + output.join(", ") + "]";
