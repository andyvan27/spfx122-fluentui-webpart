import * as React from 'react';
import { DocLib } from './DocLib';
import { WebPartContext } from '@microsoft/sp-webpart-base';
// import { DocLibClass } from './DocLibClass';

interface ToyListProps {
    listTitle: string;
    environmentMessage: string;
    userDisplayName: string;
    context: WebPartContext;
}

export const ToyList: React.FC<ToyListProps> = ({
    listTitle,
    environmentMessage,
    userDisplayName,
    context,
}) => {
    return (
        <>
            <div>
                <h3>Toy List</h3>
                <ul>
                    <li>List Title: {listTitle}</li>
                    <li>Env: {environmentMessage}</li>
                    <li>User: {userDisplayName}</li>
                </ul>
            </div>
            <div>
                <DocLib context={context} listTitle={listTitle} />
                {/* <DocLibClass /> */}
            </div>
        </>
    );
}