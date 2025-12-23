import * as React from 'react';
import { DocLib } from './DocLib';
import { DocLibClass } from './DocLibClass';

interface ToyListProps {
    description: string;
    environmentMessage: string;
    userDisplayName: string;
}

export const ToyList: React.FC<ToyListProps> = ({
    description,
    environmentMessage,
    userDisplayName,
}) => {
    return (
        <>
            <div>
                <h3>Toy List</h3>
                <ul>
                    <li>description: {description}</li>
                    <li>environmentMessage: {environmentMessage}</li>
                    <li>userDisplayName: {userDisplayName}</li>
                </ul>
            </div>
            <div>
                <DocLib />
                <DocLibClass />
            </div>
        </>
    );
}