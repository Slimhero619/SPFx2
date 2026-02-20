/**
 * WpWork.tsx
 * Root component that renders the AppBoilerplate Single Part App Page
 */

import * as React from 'react';
import {
  Stack,
  Spinner,
  SpinnerSize,
} from '@fluentui/react';
import AppBoilerplate from './AppBoilerplate';
import type { IWpWorkProps } from './IWpWorkProps';

export const WpWork: React.FC<IWpWorkProps> = ({
  context,
  isDarkTheme,
  userDisplayName,
}) => {
  return (
    <Stack
      styles={{
        root: {
          width: '100%',
          height: '100%',
          display: 'flex',
          flexDirection: 'column',
        },
      }}
    >
      <AppBoilerplate
        context={context}
        isDarkTheme={isDarkTheme}
        userDisplayName={userDisplayName}
      />
    </Stack>
  );
};

export default WpWork;
