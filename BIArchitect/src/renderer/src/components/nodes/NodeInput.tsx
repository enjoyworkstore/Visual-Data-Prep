import { memo, useEffect, useState } from 'react';

type NodeInputProps = {
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  className?: string;
};

const NodeInput = memo(({ value, onChange, placeholder, className }: NodeInputProps) => {
  const [val, setVal] = useState(value || '');

  useEffect(() => {
    setVal(value || '');
  }, [value]);

  return (
    <input
      type="text"
      className={`nodrag ${className || ''}`}
      placeholder={placeholder}
      value={val}
      onChange={(e) => setVal(e.target.value)}
      onBlur={() => onChange(val)}
      onKeyDown={(e) => {
        if (e.key === 'Enter') {
          e.currentTarget.blur();
        }
      }}
    />
  );
});

NodeInput.displayName = 'NodeInput';

export default NodeInput;
