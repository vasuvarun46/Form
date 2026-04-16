const strings = new Proxy(
  {},
  {
    get: (_target, prop) => String(prop),
  }
);

export default strings;
