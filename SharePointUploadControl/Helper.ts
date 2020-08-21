export const getSharePointFolderName = (input: string): string => {
  const formattedFolderName = input.replace(/[~.{}|&;$%@"?#<>+]/g, "-");
  return formattedFolderName;
};
