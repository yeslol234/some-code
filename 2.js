$("#setup").click(() => tryCatch(setup));
$("#basic-search").click(() => tryCatch(basicSearch));
$("#wildcard-search").click(() => tryCatch(wildcardSearch));

async function basicSearch() {
  await Word.run(async (context) => {
    let results = context.document.body.search("Online");
    results.load("length");

    await context.sync();

    // Let's traverse the search results... and highlight...
    for (let i = 0; i < results.items.length; i++) {
      results.items[i].font.highlightColor = "yellow";
    }

    await context.sync();
  });
}

async function wildcardSearch() {
  await Word.run(async (context) => {
    // Check out how wildcard expression are built, also use the second parameter of the search method to include search modes
    // (i.e. use wildcards).
    let results = context.document.body.search("$*.[0-9][0-9]", { matchWildcards: true });
    results.load("length");

    await context.sync();

    // Let's traverse the search results... and highlight...
    for (let i = 0; i < results.items.length; i++) {
      results.items[i].font.highlightColor = "red";
      results.items[i].font.color = "white";
    }

    await context.sync();
  });
}

async function setup() {
  await Word.run(async (context) => {
    context.document.body.clear();
    context.document.body.insertParagraph(
      "Video provides a powerful way to help you prove your point. When you click Online Video ($10,000.00), you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.",
      "Start"
    );
    context.document.body.paragraphs
      .getLast()
      .insertText(
        "To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching Online cover page, header, and sidebar. Click Insert and then choose the Online elements you want from the different Online galleries.",
        "Replace"
      );

    await context.sync();
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
