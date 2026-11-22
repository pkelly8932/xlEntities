export const arrayFormat = {
  isFormat(input) {
    return Array.isArray(input);
  },
  toJS(input) {
    return input;
  }
};
