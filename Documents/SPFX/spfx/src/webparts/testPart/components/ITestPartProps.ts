/* eslint-disable @typescript-eslint/no-explicit-any */
export interface ITestPartProps {
  context: {
    spHttpClient: any;
    pageContext: {
      web: {
        absoluteUrl: string;
      };
    };
  };
}
