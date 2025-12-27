import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
// import { DocLib } from './DocLib';
// import { DocLibPaged } from './DocLibPaged';
import { DocLibStreamPaged } from './DocLibStreamPaged';

interface ToyListProps {
    listTitle: string;
    listViewName: string;
    environmentMessage: string;
    userDisplayName: string;
    context: WebPartContext;
}

export const ToyList: React.FC<ToyListProps> = ({
    listTitle,
    listViewName,
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
                    <li>List View Name: {listViewName}</li>
                    <li>Env: {environmentMessage}</li>
                    <li>User: {userDisplayName}</li>
                </ul>
            </div>
            <div>
                {/* <DocLib context={context} listTitle={listTitle} />
                <DocLibPaged context={context} listTitle={listTitle} listViewName={listViewName} /> */}
                <DocLibStreamPaged context={context} listTitle={listTitle} listViewName={listViewName} />
            </div>
        </>
    );
}