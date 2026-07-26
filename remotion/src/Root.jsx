import React from "react";
import {Composition} from "remotion";
import {StoreAsset, getAssetConfig, storeAssetIds} from "./StoreAssets";

export const RemotionRoot = () => (
  <>
    {storeAssetIds.map((assetId) => {
      const asset = getAssetConfig(assetId);
      return (
        <Composition
          key={assetId}
          id={assetId}
          component={StoreAsset}
          width={asset.width}
          height={asset.height}
          fps={30}
          durationInFrames={1}
          defaultProps={{assetId}}
        />
      );
    })}
  </>
);

