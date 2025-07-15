import * as React from 'react';
import SignatureCanvas from 'react-signature-canvas';
import type { SignatureCanvasProps } from 'react-signature-canvas';

type Props = SignatureCanvasProps & {
  penColor?: string;
  backgroundColor?: string;
};

const WrappedSignatureCanvas = React.forwardRef<SignatureCanvas, Props>((props, ref) => {
  return <SignatureCanvas {...props} ref={ref} />;
});

export default WrappedSignatureCanvas;

