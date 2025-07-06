function main(args) {
  const name = args.name || "World";
  return {
    body: `Hello, ${name}!`
  };
}

exports.main = main;
