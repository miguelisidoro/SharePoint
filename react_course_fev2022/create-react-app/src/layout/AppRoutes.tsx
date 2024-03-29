import React from 'react';
import { Route, Routes } from "react-router-dom"
import { Component11, Component4, Component8, Component9 } from '../components';
import Component1 from '../components/component1';
import Component2 from '../components/component2';
import Component3 from '../components/component3';
import { NTFCollectionsDetailPage } from '../containers/NTFCollectionsDetailPage';
import { NTFCollectionsListPage } from '../containers/NTFCollectionsListPage';

export const PrivateRoutes = () => {
    return (
        <Routes>
            <Route path="/" element={<Component1 />} />
            <Route path="/" element={<Component2 />} />
        </Routes>
    )
}

export const PublicRoutes = () => {
    return (
        <Routes>
            <Route path="/" element={<Component3 />} />
            <Route path="/" element={<Component4 message="Hello!" />} />
        </Routes>
    )
}

export const AppRoutes = () => {
    return (
        <Routes>
            <Route path="/collections/:collectionId" element={<NTFCollectionsDetailPage />} />
            <Route path="/collections" element={<NTFCollectionsListPage />} />
            <Route path="/" element={<NTFCollectionsListPage />} />
        </Routes>
    )
}
